use ooxml_common::{depth::parse_guarded, ns::is_x_ns, zip::read_zip_string};

const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

use crate::{
    resolve_zip_path, CellRange, PivotCacheSource, PivotDataField, PivotDiagnostic,
    PivotDiagnosticReason, PivotLocation, PivotMetadataStatus, PivotPageField, PivotPartialReason,
    PivotTableMetadata, XlsxZip,
};

/// Parse pivot metadata without evaluating it. Saved worksheet cells and styles
/// remain authoritative and this module never reads pivotCacheRecords.
pub(crate) fn load_sheet_pivots(
    archive: &mut XlsxZip,
    sheet_path: &str,
) -> (Vec<PivotTableMetadata>, Vec<PivotDiagnostic>) {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return (Vec::new(), Vec::new());
    };
    let rels_path = format!("xl/{sheet_dir}/_rels/{sheet_file}.rels");
    let Ok(rels_xml) = read_zip_string(archive, &rels_path) else {
        return (Vec::new(), Vec::new());
    };
    let Ok(rels) = parse_guarded(&rels_xml) else {
        return (Vec::new(), Vec::new());
    };
    let base = format!("xl/{sheet_dir}");
    let targets: Vec<_> = rels
        .root_element()
        .children()
        .filter(|n| n.is_element())
        .filter(|n| n.tag_name().namespace() == Some(PACKAGE_REL_NS))
        .filter(|n| {
            n.attribute("Type")
                .is_some_and(|t| t.ends_with("/pivotTable"))
        })
        .filter_map(|n| n.attribute("Target"))
        .map(|target| resolve_zip_path(&base, target))
        .collect();

    let mut tables = Vec::new();
    let mut diagnostics = Vec::new();
    for part in targets {
        let xml = match read_zip_string(archive, &part) {
            Ok(xml) => xml,
            Err(_) => {
                diagnostics.push(PivotDiagnostic {
                    part,
                    reason: PivotDiagnosticReason::UnreadablePart,
                });
                continue;
            }
        };
        match parse_pivot_table(archive, &part, &xml) {
            Ok(table) => tables.push(table),
            Err(reason) => diagnostics.push(PivotDiagnostic { part, reason }),
        }
    }
    (tables, diagnostics)
}

fn parse_pivot_table(
    archive: &mut XlsxZip,
    part: &str,
    xml: &str,
) -> Result<PivotTableMetadata, PivotDiagnosticReason> {
    let doc = parse_guarded(xml).map_err(|_| PivotDiagnosticReason::MalformedXml)?;
    let root = doc.root_element();
    if root.tag_name().name() != "pivotTableDefinition" || !is_x_ns(root.tag_name().namespace()) {
        return Err(PivotDiagnosticReason::MalformedXml);
    }
    let name = root
        .attribute("name")
        .filter(|v| !v.trim().is_empty())
        .ok_or(PivotDiagnosticReason::MissingIdentity)?
        .to_string();
    let cache_id = root
        .attribute("cacheId")
        .and_then(|v| v.parse().ok())
        .ok_or(PivotDiagnosticReason::MissingIdentity)?;
    let location = child(root, "location")
        .and_then(parse_location)
        .ok_or(PivotDiagnosticReason::InvalidLocation)?;

    let (row_fields, malformed_rows) = parse_signed_fields(root, "rowFields");
    let (column_fields, malformed_columns) = parse_signed_fields(root, "colFields");
    let mut malformed_pages = false;
    let page_fields = child(root, "pageFields")
        .map(|parent| {
            parent
                .children()
                .filter(|n| {
                    n.is_element()
                        && n.tag_name().name() == "pageField"
                        && is_x_ns(n.tag_name().namespace())
                })
                .filter_map(|n| match n.attribute("fld").and_then(|v| v.parse().ok()) {
                    Some(field) => Some(PivotPageField {
                        field,
                        item: match n.attribute("item") {
                            Some(value) => match value.parse().ok() {
                                Some(item) => Some(item),
                                None => {
                                    malformed_pages = true;
                                    None
                                }
                            },
                            None => None,
                        },
                        name: n.attribute("name").map(str::to_string),
                    }),
                    None => {
                        malformed_pages = true;
                        None
                    }
                })
                .collect()
        })
        .unwrap_or_default();
    let mut malformed_data = false;
    let data_fields = child(root, "dataFields")
        .map(|parent| {
            parent
                .children()
                .filter(|n| {
                    n.is_element()
                        && n.tag_name().name() == "dataField"
                        && is_x_ns(n.tag_name().namespace())
                })
                .filter_map(|n| match n.attribute("fld").and_then(|v| v.parse().ok()) {
                    Some(field) => Some(PivotDataField {
                        field,
                        subtotal: n.attribute("subtotal").unwrap_or("sum").to_string(),
                        name: n.attribute("name").map(str::to_string),
                    }),
                    None => {
                        malformed_data = true;
                        None
                    }
                })
                .collect()
        })
        .unwrap_or_default();

    let mut reasons = unsupported_features(root);
    if malformed_rows {
        reasons.push(PivotPartialReason::MalformedField {
            field: "rowFields".into(),
        });
    }
    if malformed_columns {
        reasons.push(PivotPartialReason::MalformedField {
            field: "columnFields".into(),
        });
    }
    if malformed_pages {
        reasons.push(PivotPartialReason::MalformedField {
            field: "pageFields".into(),
        });
    }
    if malformed_data {
        reasons.push(PivotPartialReason::MalformedField {
            field: "dataFields".into(),
        });
    }
    let mut refresh_on_load = None;
    let mut cache_definition_part = None;
    let mut cache_source = None;
    let mut recorded_extension_uris = extension_uris(root);
    match pivot_cache_target(archive, part) {
        CacheLink::Missing => reasons.push(PivotPartialReason::MissingCacheRelationship),
        CacheLink::Ambiguous => reasons.push(PivotPartialReason::AmbiguousCacheRelationship),
        CacheLink::Target(target) => {
            cache_definition_part = Some(target.clone());
            match read_zip_string(archive, &target) {
                Err(_) => reasons.push(PivotPartialReason::UnreadableCacheDefinition),
                Ok(cache_xml) => match parse_guarded(&cache_xml) {
                    Err(_) => reasons.push(PivotPartialReason::MalformedCacheDefinition),
                    Ok(cache_doc) => {
                        let cache_root = cache_doc.root_element();
                        if cache_root.tag_name().name() != "pivotCacheDefinition"
                            || !is_x_ns(cache_root.tag_name().namespace())
                        {
                            reasons.push(PivotPartialReason::MalformedCacheDefinition);
                        } else {
                            match parse_bool(cache_root.attribute("refreshOnLoad")) {
                                Ok(value) => refresh_on_load = Some(value),
                                Err(()) => reasons.push(PivotPartialReason::MalformedField {
                                    field: "refreshOnLoad".into(),
                                }),
                            }
                            recorded_extension_uris.extend(extension_uris(cache_root));
                            if let Some(source) = child(cache_root, "cacheSource") {
                                let kind = source.attribute("type").unwrap_or("");
                                cache_source = match kind {
                                    "worksheet" => {
                                        let ws = child(source, "worksheetSource");
                                        Some(PivotCacheSource::Worksheet {
                                            sheet: ws
                                                .and_then(|n| n.attribute("sheet"))
                                                .map(str::to_string),
                                            reference: ws
                                                .and_then(|n| n.attribute("ref"))
                                                .map(str::to_string),
                                            name: ws
                                                .and_then(|n| n.attribute("name"))
                                                .map(str::to_string),
                                        })
                                    }
                                    "external" => Some(PivotCacheSource::External),
                                    "consolidation" => Some(PivotCacheSource::Consolidation),
                                    "scenario" => Some(PivotCacheSource::Scenario),
                                    _ => None,
                                };
                                if kind != "worksheet" {
                                    reasons.push(PivotPartialReason::UnsupportedCacheSource {
                                        source_type: kind.to_string(),
                                    });
                                }
                            }
                            for feature in [
                                "tupleCache",
                                "calculatedItems",
                                "calculatedMembers",
                                "dimensions",
                                "measureGroups",
                                "maps",
                            ] {
                                if child(cache_root, feature).is_some() {
                                    reasons.push(PivotPartialReason::UnsupportedSemanticFeature {
                                        feature: feature.to_string(),
                                    });
                                }
                            }
                            if cache_root.descendants().any(|node| {
                                node.is_element()
                                    && node.tag_name().name() == "fieldGroup"
                                    && is_x_ns(node.tag_name().namespace())
                            }) {
                                reasons.push(PivotPartialReason::UnsupportedSemanticFeature {
                                    feature: "fieldGroup".into(),
                                });
                            }
                        }
                    }
                },
            }
        }
    }

    let status = if reasons.is_empty() {
        PivotMetadataStatus::Complete
    } else {
        PivotMetadataStatus::Partial { reasons }
    };
    Ok(PivotTableMetadata {
        name,
        cache_id,
        location,
        row_fields,
        column_fields,
        page_fields,
        data_fields,
        refresh_on_load,
        cache_definition_part,
        cache_source,
        status,
        extension_uris: recorded_extension_uris,
    })
}

enum CacheLink {
    Missing,
    Ambiguous,
    Target(String),
}

fn pivot_cache_target(archive: &mut XlsxZip, pivot_part: &str) -> CacheLink {
    let Some((dir, file)) = pivot_part.rsplit_once('/') else {
        return CacheLink::Missing;
    };
    let Ok(xml) = read_zip_string(archive, &format!("{dir}/_rels/{file}.rels")) else {
        return CacheLink::Missing;
    };
    let Ok(doc) = parse_guarded(&xml) else {
        return CacheLink::Missing;
    };
    let targets: Vec<_> = doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
        .filter(|n| n.tag_name().namespace() == Some(PACKAGE_REL_NS))
        .filter(|n| {
            n.attribute("Type")
                .is_some_and(|t| t.ends_with("/pivotCacheDefinition"))
        })
        .filter_map(|n| n.attribute("Target"))
        .collect();
    match targets.as_slice() {
        [target] => CacheLink::Target(resolve_zip_path(dir, target)),
        [] => CacheLink::Missing,
        _ => CacheLink::Ambiguous,
    }
}

fn parse_location(node: roxmltree::Node<'_, '_>) -> Option<PivotLocation> {
    Some(PivotLocation {
        range: parse_a1_range(node.attribute("ref")?)?,
        first_header_row: node.attribute("firstHeaderRow")?.parse().ok()?,
        first_data_row: node.attribute("firstDataRow")?.parse().ok()?,
        first_data_col: node.attribute("firstDataCol")?.parse().ok()?,
    })
}

fn parse_a1_range(value: &str) -> Option<CellRange> {
    let (first, last) = value.split_once(':').unwrap_or((value, value));
    let (left, top) = parse_a1_cell(first)?;
    let (right, bottom) = parse_a1_cell(last)?;
    (top <= bottom && left <= right).then_some(CellRange {
        top,
        left,
        bottom,
        right,
    })
}

fn parse_a1_cell(value: &str) -> Option<(u32, u32)> {
    let value = value.replace('$', "");
    let split = value.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = value.split_at(split);
    if letters.is_empty()
        || digits.is_empty()
        || !letters.chars().all(|c| c.is_ascii_alphabetic())
        || !digits.chars().all(|c| c.is_ascii_digit())
    {
        return None;
    }
    let col = letters.chars().try_fold(0u32, |acc, c| {
        acc.checked_mul(26)?
            .checked_add((c.to_ascii_uppercase() as u8 - b'A' + 1) as u32)
    })?;
    let row = digits.parse().ok()?;
    (col > 0 && row > 0).then_some((col, row))
}

fn parse_signed_fields(root: roxmltree::Node<'_, '_>, parent_name: &str) -> (Vec<i32>, bool) {
    let mut malformed = false;
    let fields = child(root, parent_name)
        .map(|parent| {
            parent
                .children()
                .filter(|n| {
                    n.is_element()
                        && n.tag_name().name() == "field"
                        && is_x_ns(n.tag_name().namespace())
                })
                .filter_map(|n| match n.attribute("x").and_then(|v| v.parse().ok()) {
                    Some(value) => Some(value),
                    None => {
                        malformed = true;
                        None
                    }
                })
                .collect()
        })
        .unwrap_or_default();
    (fields, malformed)
}

fn unsupported_features(root: roxmltree::Node<'_, '_>) -> Vec<PivotPartialReason> {
    [
        "pivotHierarchies",
        "filters",
        "rowHierarchiesUsage",
        "colHierarchiesUsage",
    ]
    .into_iter()
    .filter(|feature| child(root, feature).is_some())
    .map(|feature| PivotPartialReason::UnsupportedSemanticFeature {
        feature: feature.to_string(),
    })
    .collect()
}

fn extension_uris(root: roxmltree::Node<'_, '_>) -> Vec<String> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "ext")
        .filter_map(|n| n.attribute("uri"))
        .map(str::to_string)
        .collect()
}

fn child<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children().find(|child| {
        child.is_element()
            && child.tag_name().name() == name
            && is_x_ns(child.tag_name().namespace())
    })
}

fn parse_bool(value: Option<&str>) -> Result<bool, ()> {
    match value {
        None | Some("0" | "false") => Ok(false),
        Some("1" | "true") => Ok(true),
        Some(_) => Err(()),
    }
}
