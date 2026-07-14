import { useState } from "react";
import ColorSwatch from "./ColorSwatch.jsx";

// One editable row per distinct color. Labels are saved per color (the
// backend attaches a label to a color, so every cell of that color shares
// it). Explicit "Save" button per row — nothing auto-saves on blur.
function LabelRow({ entry, onLabel }) {
  const initial = entry.human_label || "";
  const [value, setValue] = useState(initial);
  const [saving, setSaving] = useState(false);
  const dirty = value.trim() !== initial;

  const save = async () => {
    setSaving(true);
    try {
      await onLabel(entry.key, value.trim());
    } finally {
      setSaving(false);
    }
  };

  return (
    <tr>
      <td>
        <ColorSwatch
          color={{
            resolved: entry.resolved,
            hex: entry.resolved ? entry.hex_or_theme_ref.replace("#", "") : null,
            theme_ref: !entry.resolved ? entry.hex_or_theme_ref : null,
            reason_unresolved: entry.reason_unresolved,
          }}
        />
      </td>
      <td>{entry.cell_count}</td>
      <td className="mono">{(entry.example_refs || []).join(", ")}</td>
      <td>
        <input
          type="text"
          value={value}
          placeholder="label"
          onChange={(e) => setValue(e.target.value)}
        />
      </td>
      <td>
        <button
          className="secondary"
          onClick={save}
          disabled={saving || !dirty}
        >
          {saving ? "Saving…" : "Save"}
        </button>
      </td>
    </tr>
  );
}

export default function ColorInventoryTable({ inventory, onLabel }) {
  if (!inventory || inventory.length === 0) {
    return <p className="note">No colors were found in this sheet.</p>;
  }

  return (
    <table>
      <thead>
        <tr>
          <th>Color</th>
          <th>Cells</th>
          <th>Example cells</th>
          <th>Label</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        {inventory.map((entry) => (
          <LabelRow key={entry.key} entry={entry} onLabel={onLabel} />
        ))}
      </tbody>
    </table>
  );
}
