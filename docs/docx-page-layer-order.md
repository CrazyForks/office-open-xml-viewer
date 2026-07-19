# DOCX page-layer ordering evidence

This note is the decision table for Series B3 of issue #1037. It separates
normative WordprocessingML anchor ordering from cross-story compatibility
behavior so the renderer does not turn an observation into a specification
claim.

## Normative anchor rule

ECMA-376 Part 1 §20.4.2.3 (`wp:anchor`) defines the relative z-order for
floating DrawingML objects:

- every object with `behindDoc=false` is above every object with
  `behindDoc=true`;
- among objects with the same `behindDoc` value, a larger `relativeHeight`
  value is higher in the z-order.

The retained layout uses source order and then stable encounter order only to
make equal-`relativeHeight` ties deterministic. Those tie breakers do not
override either normative rule.

## Story decision table

| Owner | Equivalent stacking context | Anchors materialized in `PageLayers.paintOrder` | Cross-story decision | Evidence |
| --- | --- | --- | --- | --- |
| Body | All consecutive body roots on one physical page | Behind anchors, body ink roots, front anchors | Preserve the established page-root position relative to header, notes, and footer | §20.4.2.3 for within-context order; synthetic regression coverage for preservation |
| Header | The selected header story for the physical page | Behind anchors, header ink roots, front anchors | Keep the complete header context at its retained root position | §20.4.2.3 for within-context order; synthetic regression coverage for preservation |
| Footer | The selected footer story for the physical page | Behind anchors, footer ink roots, front anchors | Keep the complete footer context at its retained root position | §20.4.2.3 for within-context order; synthetic regression coverage for preservation |
| Footnote/endnote | Each retained notes-layer run | Behind anchors, note ink roots, front anchors | Keep the notes context at its retained root position | §20.4.2.3 for within-context order; synthetic regression coverage for preservation |
| Text box | The owner drawing and its complete text-box story | The owner is one page entry; descendant text-box anchors remain local and are not page-queued again | Atomic with the owner drawing | Retained ownership invariant and synthetic regression coverage |

ECMA-376 does not define one total z-order across the body, header, footer,
note, and text-box stories. The current B3 implementation therefore preserves
the pre-refactor cross-story root order while moving the ordering authority from
paint-time callbacks into immutable layout data.

## Compatibility-observation status

The reproducible observation matrix consists of pairwise overlapping anchors in
header, body, footnote, and footer stories, with `relativeHeight` swapped for
each pair. It is intentionally synthetic and contains no private document data.

As of this change, desktop Word on the fresh test machine can open the generated
documents, but AppleEvent export is blocked by the machine's automation
permission state; Word for the web also requires an authenticated session.
Consequently, no cross-story result is labeled as an Office observation yet.
When the synthetic matrix can be rendered by Word, its application version,
platform, fixture generator, and observed pixels must be recorded before any
cross-story compatibility rule changes.

## Paint-plan invariants

- `PageLayers.roots` owns top-level story graph/composition order.
- `PageLayers.paintOrder` is the only final paint-order authority.
- Every top-level root is represented by at least one paint entry.
- Every page-materialized anchored drawing appears exactly once.
- A drawing entry retains only plain-data affine, clip, coordinate-space, and
  owner facts; it cannot retain a Canvas closure.
- Table, cell, resolved floating-table, paragraph, and section transforms are
  replayed exactly once.
- A drawing and the text boxes it owns paint atomically.
- Descendant text-box anchors remain in the owner's local stacking context.
- Canvas paint performs no graph discovery, z-order sorting, measurement, or
  layout mutation.
