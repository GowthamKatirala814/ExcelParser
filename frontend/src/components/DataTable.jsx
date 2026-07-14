import ColorSwatch from "./ColorSwatch.jsx";

// Read-only per-cell table. The label column reflects the label saved for
// each cell's color in the Colors table above (labels are per color).
export default function DataTable({ rows }) {
  if (!rows || rows.length === 0) {
    return <p className="note">No cells found in this sheet.</p>;
  }

  return (
    <div className="table-scroll">
      <table>
        <thead>
          <tr>
            <th>Cell</th>
            <th>Value</th>
            <th>Color</th>
            <th>Label</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.cell_ref}>
              <td className="mono">{row.cell_ref}</td>
              <td>
                {row.value === null || row.value === undefined ? (
                  <span className="swatch-none">blank</span>
                ) : (
                  String(row.value)
                )}
              </td>
              <td>
                <ColorSwatch color={row.color} />
              </td>
              <td>{row.human_label || ""}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
