//! Isolated Office compatibility rules for shared DrawingML transforms.

/// Resolve the group scale applied to a retained pre-rotation leaf frame.
///
/// ECMA-376 Part 1 Annex L §L.4.7.4 describes scale accumulation on authored
/// horizontal and vertical axes. [MS-OE376] §2.1.1360 additionally defines the
/// Office group scale as `ext / chExt`. Office-produced reference output shows
/// that, for an exact odd quarter-turn leaf, Word applies that ratio on the
/// post-rotation page axes. Exchanging the retained frame axes represents that
/// case exactly; arbitrary angles remain on the normative Annex L path because
/// an axis-aligned rectangle plus rotation cannot represent their skew.
pub(crate) fn office_group_scale_for_leaf_rotation(
    scale_x: f64,
    scale_y: f64,
    rotation_degrees: f64,
) -> (f64, f64) {
    let quarter_turns = (rotation_degrees / 90.0).round();
    let exact_quarter_turn = (rotation_degrees - quarter_turns * 90.0).abs() < 1e-9;
    if exact_quarter_turn && (quarter_turns as i64).rem_euclid(2) == 1 {
        (scale_y, scale_x)
    } else {
        (scale_x, scale_y)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn swaps_only_exact_odd_quarter_turns() {
        assert_eq!(
            office_group_scale_for_leaf_rotation(2.0, 3.0, 90.0),
            (3.0, 2.0)
        );
        assert_eq!(
            office_group_scale_for_leaf_rotation(2.0, 3.0, 270.0),
            (3.0, 2.0)
        );
        assert_eq!(
            office_group_scale_for_leaf_rotation(2.0, 3.0, 180.0),
            (2.0, 3.0)
        );
        assert_eq!(
            office_group_scale_for_leaf_rotation(2.0, 3.0, 45.0),
            (2.0, 3.0)
        );
    }
}
