# grid

Canvas geometry/rendering, input, DOM editing and accessible grid projection.

`Worksheet.tsx` projects immutable sheet snapshots. `grid-geometry.ts` supplies the axis coordinates used by the canvas, hit tests and scroll reveal. `grid-selection.ts` owns pure selection and navigation functions. The React worksheet manages view state and the bounded semantic projection. Authored edits belong to the workbook session.
