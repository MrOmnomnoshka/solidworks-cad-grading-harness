# Scoring Notes

## Scoring approach

The harness evaluates candidate SolidWorks parts using measured geometry rather than feature or body names. It assigns bodies to functional roles, compares their mass properties and centroids against the baseline, evaluates the widened stance and button-cluster placement, checks the asymmetric shoulder-marking signature, and estimates the presence of modeled 3D button-symbol details.

The overall score is a weighted sum of continuous component scores, with a maximum of 4.0 points.

### Intentional centre-position penalty

The centre-position component deliberately retains a small penalty when measured centre positions visibly deviate from the expected mirrored targets. This includes the supplied reference solution: manual inspection in SolidWorks confirms that the displacement is visible in the actual geometry, rather than being merely numerical noise.

The same geometric rule is applied consistently to every candidate, without special-casing the reference by filename. Consequently, the reference solution scores 3.965/4.000 instead of full marks. This is a deliberate choice to preserve a stricter geometric interpretation of the result.

## Dataset observations

The supplied CAD examples do not always isolate a single advertised failure. Some files contain additional geometric changes or rebuild errors, and in some cases the measured geometry differs from what the filename or description suggests. The harness therefore reports observed geometric behavior rather than inferring a model's properties from its filename.

The button-symbol component measures 3D face-level detail. It does not perform OCR or assess the identity or chirality of individual symbols. Visible flat or sketch-level markings may receive less credit than modeled embossed or recessed geometry.

The shoulder-marking component uses the asymmetric face-area signature of the R glyph as a proxy for detecting an incorrect mirror operation.

## Handling incomplete checks

Pairwise interference checks are performed on the mapped reference-role bodies. If role matching is incomplete or a Boolean operation fails, the interference result is marked unavailable and a warning is emitted. The current scoring policy assigns zero points to this component when it cannot be fully verified.

Small incidental intersections (V=2.17 mm³) at mounting interfaces are treated as minor imperfections rather than automatically being classified as major collisions.