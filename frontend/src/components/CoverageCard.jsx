// Plain-text coverage summary for one sheet. No card, no colors.
const fmt = (n) => (n ?? 0).toLocaleString();

export default function CoverageCard({ summary }) {
  if (!summary) return null;
  const total = summary.total_cells || 0;

  return (
    <div>
      <p className="status-line">
        {fmt(summary.colored_cells)} of {fmt(total)} cells had a resolved color,{" "}
        {fmt(summary.unresolved_cells)} unresolved, {fmt(summary.blank_cells)} blank
        {summary.ambiguous_cells > 0
          ? `, ${fmt(summary.ambiguous_cells)} ambiguous (multiple rules matched)`
          : ""}
        .
      </p>
      {summary.processing_time_seconds != null && (
        <p className="note">Processed in {summary.processing_time_seconds.toFixed(3)}s.</p>
      )}
    </div>
  );
}
