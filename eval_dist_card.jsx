import { PieChart, Pie, Cell, Tooltip } from "recharts";

const data = [
  { name: "M7", value: 15, color: "#4F81BD" },
  { name: "Parentis", value: 5, color: "#2C3E8C" },
  { name: "Britelift", value: 18, color: "#C0392B" },
  { name: "RideX", value: 18, color: "#8E44AD" },
];

const percentages = { M7: "26.8%", Parentis: "8.9%", Britelift: "32.1%", RideX: "32.1%" };
const RADIAN = Math.PI / 180;

const renderCustomLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, value, index }) => {
  const cosA = Math.cos(-midAngle * RADIAN);
  const sinA = Math.sin(-midAngle * RADIAN);
  const midR = (innerRadius + outerRadius) / 2;
  const color = data[index].color;
  const ix = cx + midR * cosA;
  const iy = cy + midR * sinA;
  const lx1 = cx + (outerRadius + 4) * cosA;
  const ly1 = cy + (outerRadius + 4) * sinA;
  const lx2 = cx + (outerRadius + 16) * cosA;
  const ly2 = cy + (outerRadius + 16) * sinA;
  const tickDir = cosA >= 0 ? 1 : -1;
  const lx3 = lx2 + tickDir * 8;

  return (
    <g>
      <text x={ix} y={iy} textAnchor="middle" dominantBaseline="central"
        fill="white" fontSize={13} fontWeight="bold" style={{ pointerEvents: "none" }}>
        {value}
      </text>
      <line x1={lx1} y1={ly1} x2={lx2} y2={ly2} stroke={color} strokeWidth={1.5} />
      <line x1={lx2} y1={ly2} x2={lx3} y2={ly2} stroke={color} strokeWidth={1.5} />
      <text x={lx3 + tickDir * 3} y={ly2} textAnchor={cosA >= 0 ? "start" : "end"}
        dominantBaseline="central" fill={color} fontSize={10} fontWeight="bold"
        style={{ pointerEvents: "none" }}>
        {data[index].name}
      </text>
    </g>
  );
};

const CustomTooltip = ({ active, payload }) => {
  if (!active || !payload?.length) return null;
  const e = payload[0];
  const total = data.reduce((s, d) => s + d.value, 0);
  const pct = (e.value / total * 100).toFixed(1);
  return (
    <div style={{
      background: "#fff", border: "1px solid #E2E8F0", borderRadius: 8,
      padding: "8px 12px", boxShadow: "0 2px 8px rgba(0,0,0,.10)", fontSize: 12,
    }}>
      <div style={{ fontWeight: 700, color: e.payload.color, marginBottom: 2 }}>{e.name}</div>
      <div style={{ color: "#475569" }}>{e.value} evaluations</div>
      <div style={{ color: "#94A3B8", marginTop: 1 }}>{pct}% of total</div>
    </div>
  );
};

export default function App() {
  return (
    <div style={{
      width: 340, background: "#fff", border: "1px solid #E2E8F0",
      borderRadius: 12, padding: "16px 16px 0", boxSizing: "border-box",
      fontFamily: "system-ui, sans-serif",
    }}>
      <div style={{ marginBottom: 2 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: "#1E293B" }}>Evaluation distribution</div>
        <div style={{ fontSize: 11, color: "#94A3B8", marginTop: 2 }}>56 evaluations across all accounts</div>
      </div>

      <PieChart width={340} height={300}>
        <Pie data={data} cx={170} cy={150} innerRadius={62} outerRadius={100}
          startAngle={90} endAngle={-270} dataKey="value"
          labelLine={false} label={renderCustomLabel}>
          {data.map((entry, i) => (
            <Cell key={i} fill={entry.color} stroke="#fff" strokeWidth={2} />
          ))}
        </Pie>
        <Tooltip content={<CustomTooltip />} />
      </PieChart>

      <div style={{
        borderTop: "1px solid #F1F5F9", paddingTop: 10, paddingBottom: 12,
        display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px 8px",
      }}>
        {data.map((d) => (
          <div key={d.name} style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 7, height: 7, borderRadius: "50%", background: d.color, flexShrink: 0 }} />
              <span style={{ fontSize: 11, color: "#475569" }}>{d.name}</span>
            </div>
            <span style={{ fontSize: 11, fontWeight: 700, color: d.color }}>{percentages[d.name]}</span>
          </div>
        ))}
      </div>
    </div>
  );
}
