// A small colored square next to its hex code. Never tints a whole row, so
// cell text stays readable. Unresolved colors show an empty bordered square
// plus a short text note (no gradients or hatching).
export default function ColorSwatch({ color }) {
  if (!color || (!color.resolved && !color.theme_ref && !color.hex)) {
    return <span className="swatch-none">none</span>;
  }

  if (color.resolved && color.hex) {
    return (
      <span className="swatch-wrap">
        <span className="swatch" style={{ backgroundColor: `#${color.hex}` }} />
        <span className="mono">#{color.hex}</span>
      </span>
    );
  }

  return (
    <span className="swatch-wrap" title={color.reason_unresolved || ""}>
      <span className="swatch" style={{ backgroundColor: "#ffffff" }} />
      <span className="mono">{color.theme_ref || "unresolved"}</span>
    </span>
  );
}
