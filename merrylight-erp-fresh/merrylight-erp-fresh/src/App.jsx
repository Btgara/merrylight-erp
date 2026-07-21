import React, { useState, useEffect, useMemo } from "react";
import {
  ResponsiveContainer, AreaChart, Area, BarChart, Bar, LineChart, Line,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, Cell,
} from "recharts";
import * as XLSX from "xlsx";



/* ================= SEED DATA (extracted from Daily_Processed_Fruit_Data_2026.xlsm) ================ */
// PRODUCTION (web-primary): batch records live entirely in the app's shared
// team storage, loaded via the Batch upload page. Season history is imported
// once from Merrylight_Season_History.xlsx. No batch data is embedded here.
// WORKSHOP COPY: contains 5 sample batch records so the artifact stays under
// Claude's preview size limit. The deployed app (merrylight-pwa) carries the
// full dataset, injected at build time from merrylight-seed-full.txt.
const SEED = (() => {
  const fields = ["date", "shift", "week", "batch", "block", "variety", "section", "reaped", "processed", "bw", "kgh", "a12", "a14", "a18", "a22", "s12", "s14", "s18", "s22", "tears", "redBums", "greenBums", "mold", "stemPull", "shrivell", "scar", "softPoint", "soft", "under12", "subBloom"];
  const S = { 0: 1, 1: 1, 3: 1, 4: 1, 5: 1, 6: 1 }; // string-typed columns
  const raw = '2026-04-28|DAY|18|MER0016|C09|ARANA|COMPLEX||1475.5|3.79|600|0|4.2|177.1|667.7|0.1|21|551.3|0|11.1|10.3|5.1|0.5|0.5|5.8|3.6|11.8|4.7|0|0.7~2026-04-29|DAY|18|MER0017|C15|ARANA|COMPLEX||1679.9|3.8|800|0.1|11.5|229.4|653.8|0.1|48.6|676.6|0|14.4|5.6|5.5|0.8|0.6|5.8|8|12.9|5.5|0|0.6~2026-04-29|DAY|18|MER0019|C10|ARANA|COMPLEX||1569.2|3.8|780|0|6.4|214.4|622.1|0|31.7|650.7|0|5.8|4.2|4.8|0|0.6|14|7.4|0.3|3|0|0.7~2026-04-30|DAY|18|MER0020|C16|ARANA|COMPLEX||2114.9|3.7|820|0|9.4|311.1|848.3|0|41.2|837.3|0|6.4|9.6|4.8|0|1|24.2|11.4|4.5|4.1|0|1.4~2026-04-30|DAY|18|MER0021|C21|ARANA|COMPLEX||2500.3|3.5|800|0|15|385.4|811.2|0|74.6|1131.2|0|8.8|11.3|7.2|0|1.4|29.7|17.2|2.9|2.7|0|1.7';
  const rows = raw ? raw.split('~').map((line) => line.split('|').map((v, i) => (v === '' ? null : S[i] ? v : +v))) : [];
  return { fields, rows };
})();

const LOGO = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAMAAACdt4HsAAAB/lBMVEUdWzBkmVnf5ekPICFMZlwxZ6UubFIWcF9mjq5gl1OQmZ+uyd2XtGAgdDpTa1hbksZdboxhml+IuTtil1f//wB/fwB//38jLWQndUVdlDyNtlSRr9lGdjqizVGYzZGdy06Su1mVxTtwm8udyVDh6KUUGGkdajtlnbSns5GVx43/AP/Y6d/o5ucTIBcleD09hEQ5kMxJbEtCaEp2YZ1/f/9//wCesJ+SqdSuzJguK10qKpQkP4AAAP8hb0Y8hj4/hVE9glc8h74A//9GfsBims14w3h///+KpVqsurWJt9y036qt0eqjwOH/AADM7mbK14T99fkAAAAkLEYrNE8qOWw3RnBJWolLVnIUGSwiKDscJk1Sdq0bIzVRZpNyeIo2VY8uSoZQZW9xmMh/f3+KqNIxaU1qhrV5gpGFiZWnqbLFxstJaqdkaXkvWU85WW0tdzf///8zRU10ptKUmqOytrvK1+YAfwAA/wAoWDBHaVdNiDlleaiUtdiyu8mrx+YRGUcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUgWX0AAAAgHRSTlP7+f7++//9Ev2l+v4dpRb5/hH5WwECAv9X8+/5+933pFHyZmIRGgkUFq0BEXsTYmT/XqwGAgJqDUlHDP8BkLgs//8B/4b/AqCqbxgypAEPdYYA//7+/v7+///////+//7+/f8C/v7+///+/v///v75Af///v7+AgH+/vn///7+/8Z/JYAAAAb3SURBVHja3Zdnd9vGEobRCZKSSFmyVS3ZjmM79k3PTXJ7TW4vAHax6I0gQYBNbCIptr+eASVbpEQ5zvG3jHQEgkfz7DtlFwNKek+jfr6As/cDbEtS6T0ApcGLv21Iv3wPgPS/h4ffroliu/ROgF1p4yFNHz6VSr/9/qtMJvPVJWlQelcF29Kjw+MWfXr24vU3z/98mdWd7duy1gAGUpZuHR8cbkhn//rHX588efINRHUmvfo6S3909k45OMu2KOrg4aNX/z4KxVqtUHjyXJL+maVp+uTpLQmrgEGpBJGebVLD4cVh7mg2y+fDyjjoFJ4/zQqNMn3yrfTgxxU8K21S1PD4cBJOZ/0+E4ZHlYqYLVNDjj756O2ArdReSdI+BQTx6GgWBOAbhhXbLlJUgz4BwFtysCttZ4+p4y8+36Apih2Fk7GtqmYnnIQV13ObFAg4ye6W3qqApodgF3SZYp1Ox2DnhhrOFgCv2RBo+rudtMvvBGxD/ajUGoRqV4IO0vYMVMnPQtP7DXLLZSjD6a0Wp1Za+AXNDVkWEGyjInYQkueGG0xF20MIC6CAPny4cSON1Oom2EklAMIipjnWxjb2sD223aaPcBkAnH8PCLt3AdIs0vRxSrDI2DZRsWKOVbVmN4YcQiPhj5RG/Hunq4SVEF5mL3zQSQ0pVhm7plmriA7yig2KKiPcbrQ1TSOjx6criVwC3Jd2LkaqMqI5asjqjurypqmiPdS2KBoAlsYCQFO5xx8sS6CW/LcOOFvRNB2y2OAi1XNdHgFAYVkuqkR2UwEEK9sHkIb7awAlCQLAmq40hLKYz+dNDyMX7e0hq93NM0zeDDpiU9NY7EMQD24D7kvbEIDcLguCKE7y+Q7P+6MyZciGP5kxTMGdoyLfKaZB0I+XavkGcA4Z8HFbEChxEuaZfsDzvADZtJpBOGOmSMZghQLRWMW/98F1DNchbB2UVVBfzvXBpv1JEPBekxu2a7WjPJPHkFHTLnxWaLMgAbJwE7Al7bR8kuqf9G2kzO0+MyuYLrawa6aC3L+YZmdc+BPPWyzmQMLGDUBJ+rRFoNfEIOzL8z1jLjKhWXM9Gdnj8ZTJ867YCdxCgefJZQwPbobw5aZAGhw3LUxsDOGjYh4EeGklVHUcTALR5M1OYPK8qVmEPv0Gjs5lwEDaapVVXSOTYOKan/H8kSoGpukihFTVU83JLAwK4aRjuoKpsyStw+4K4FzKtASiW/6s0BGdzmw29oIwqNiwC23b8+xKGIhwNAWmh2wlTcLG62ffFSADgCaWreKsU+DNihNVRNGs2SoyFAwdCTuq5tpBRQVJts6i0RrA/hUg4M2g4lVAf80j2JgbCkJuza7B0hUbp4osDbL46HUW34TwyQJAjmp8za3ValD1GiRgrhmKsufxLmTTU7Fs7KlEWwe4UiBjyDOcoJ6nui6GNgaAgWRP9TBGsmzATVHR1iuAJGJdUUAnnF8eMhAyUoAlgwSkQhsrcA9dQTRLw/6tHAAAjlJZ0ZRxqhyjuTFn53NWkzXQoBCEsbEgET3dkLerMJCeCZEqK/U2nGXYzuWI7DuOwvkWR4iCfYJlRDCAdV3n2kX68GYfgP2d6WoR4+R63Shmqn2SdOOkyshxNXEMLu8U83GiTtt6L5cwSZL7+OZmOpN+wTDtHlOMc0TJTYsJnla7sZx0u0pS1+TESaqGQxjSZupGLx5xH64B9Jgk13dySVfuMkxsTadMla0zxEoIa/WqCYliwkx7DMfG8cXOmyFsCdDNMaTv9HIKiXuwUFKvM0qbUXSmSuqJk89xjMNwCkO0Xo97Kv3nFuD/vbKjMFE16eWcyKo6DpEhmljRu72eE/skmcYk4XRg9HufS2tC+DqiBLVLoFjQMYqVVt0ydEPW4aNhGZrMGZYiQ6uQnPPh9Rh5XYWPf/U7qokNcFeUqz9w0dMf3YLqWRb8WmkXXXzx68HaJ1OGEmwMnrITRXVYJ3Kcbr0KVueqTq5e9TVd1xSbpv+79GRZAjyDyaShput2o2jqRPVuVJ1W41wcV6GiMYkd0CCro4Od5Tl4RYG0STXVhYK6X/cJSKg7jlx3It+JqvARogD/7MocTa0MiJkF4SoJl5dLU6z0Bvz9g+zWYHD3lJahFgT52jU1/dIsRfUvstvS4O4x7zzV0FDxGxErFGyPUv/7b5sTz6WtlJCKWI4g1QDLg//WDf9bg+a5VNpPw0hVLHnLiqqOjmELwD/8yKQKAUIYQsNWOYyuOgpzsDq3CH/wDsM2rLEPw25ZaBJ1YcT3hdZFdke6Kf+uaT1VmdkERqucmtBqtX4Pwa/1v+OlK5M+rzP7n2wubP/TP3y5eID/lNe+88zl9eXLq8f/+U9/8cxkzl/DMj/rd+cfANueuBb8+jW0AAAAAElFTkSuQmCC";

/* ================= NATIVE BAR CHART ================= */
// Supports single, grouped, and stacked series; negative values; per-value colors.
function niceStep(rough) {
  const mag = Math.pow(10, Math.floor(Math.log10(rough || 1)));
  for (const m of [1, 2, 2.5, 5, 10]) if (rough <= m * mag) return m * mag;
  return 10 * mag;
}
function SBarChart({ data, xKey, series, height = 240, unit = "", fmt, angled = false, stacked = false }) {
  if (!data || !data.length) return <p className="muted">No data to chart.</p>;
  const W = 720, H = height, padL = 48, padR = 10, padT = 12;
  const padB = angled ? 58 : 26;
  const F = fmt || ((v) => (Math.abs(v) >= 1000 ? v.toLocaleString("en") : +v.toFixed(2)) + unit);

  const vals = [];
  for (const d of data) {
    if (stacked) vals.push(series.reduce((t, sr) => t + Math.max(0, n0(d[sr.key])), 0));
    else series.forEach((sr) => vals.push(n0(d[sr.key])));
  }
  let max = Math.max(0, ...vals), min = Math.min(0, ...vals);
  if (max === min) max = min + 1;
  const step = niceStep((max - min) / 4);
  max = Math.ceil(max / step) * step;
  min = Math.floor(min / step) * step;
  const span = max - min || 1;
  const y = (v) => padT + ((max - v) / span) * (H - padT - padB);

  const ticks = [];
  for (let t = min; t <= max + 1e-9; t += step) ticks.push(+t.toFixed(6));

  const band = (W - padL - padR) / data.length;
  const groupW = band * 0.62;
  const subW = stacked ? groupW * 0.85 : groupW / series.length;

  return (
    <div>
      <svg viewBox={"0 0 " + W + " " + H} width="100%" style={{ display: "block" }} role="img">
        {ticks.map((t) => (
          <g key={t}>
            <line x1={padL} x2={W - padR} y1={y(t)} y2={y(t)} stroke={t === 0 ? "#C9CADB" : "#E7E8F2"} />
            <text x={padL - 6} y={y(t) + 3.5} textAnchor="end" fontSize="10" fill="#6E6D85">{F(t)}</text>
          </g>
        ))}
        {data.map((d, i) => {
          const x0 = padL + i * band + (band - groupW) / 2;
          let stackPos = 0;
          return (
            <g key={i}>
              {series.map((sr, si) => {
                const v = n0(d[sr.key]);
                let rx, ry, rh;
                if (stacked) {
                  const vv = Math.max(0, v);
                  rx = x0 + (groupW - subW) / 2;
                  ry = y(stackPos + vv);
                  rh = y(stackPos) - y(stackPos + vv);
                  stackPos += vv;
                } else {
                  rx = x0 + si * subW;
                  ry = v >= 0 ? y(v) : y(0);
                  rh = Math.abs(y(v) - y(0));
                }
                const color = sr.colorFor ? sr.colorFor(v) : sr.color;
                return (
                  <rect key={sr.key} x={rx} y={ry} width={Math.max(1, subW - (stacked ? 0 : 1.5))} height={Math.max(0, rh)} fill={color} rx="1.5">
                    <title>{(d[xKey] || "") + " \u00b7 " + (sr.label || sr.key) + ": " + F(v)}</title>
                  </rect>
                );
              })}
              {angled ? (
                <text x={x0 + groupW / 2} y={H - padB + 12} fontSize="9" fill="#232047"
                  textAnchor="end" transform={"rotate(-35 " + (x0 + groupW / 2) + " " + (H - padB + 12) + ")"}>{d[xKey]}</text>
              ) : (
                <text x={x0 + groupW / 2} y={H - 8} fontSize="11" fill="#232047" textAnchor="middle">{d[xKey]}</text>
              )}
            </g>
          );
        })}
      </svg>
      {series.length > 1 && (
        <div className="legend" style={{ justifyContent: "center", display: "flex", gap: 16, fontSize: 11.5, color: "var(--muted)", marginTop: 6 }}>
          {series.map((sr) => (
            <span key={sr.key}><i className="dot" style={{ background: sr.color }} />{sr.label || sr.key}</span>
          ))}
        </div>
      )}
    </div>
  );
}

/* ================= HELPERS ================= */
const F = SEED.fields;
const idx = Object.fromEntries(F.map((f, i) => [f, i]));
const num = (v) => (v == null || v === "" || isNaN(+v) ? null : +v);
const n0 = (v) => (num(v) == null ? 0 : +v);

function rowToRec(row, i) {
  const r = {};
  F.forEach((f, j) => (r[f] = row[j]));
  r.id = "seed-" + i;
  r.source = "workbook";
  return r;
}

const DEFECTS = [
  ["tears", "Tears"], ["redBums", "Red bums"], ["greenBums", "Green bums"],
  ["mold", "Mold"], ["stemPull", "Stem pull"], ["shrivell", "Shrivel"],
  ["scar", "Scar"], ["softPoint", "Soft point"], ["soft", "Soft"],
  ["under12", "<12mm"], ["subBloom", "Sub bloom"],
];

// Groupings taken from the workbook's own formulas
const GROUPS = {
  "Reaping issues": ["tears", "softPoint"],
  "Under ripe": ["redBums", "greenBums"],
  "Over ripe": ["mold", "shrivell", "soft"],
  "Other": ["stemPull", "scar", "subBloom"],
};

function derive(r) {
  const p = num(r.processed);
  const reaped = num(r.reaped);
  const air = n0(r.a12) + n0(r.a14) + n0(r.a18) + n0(r.a22);
  const sea = n0(r.s12) + n0(r.s14) + n0(r.s18) + n0(r.s22);
  const sz = {
    k12: n0(r.a12) + n0(r.s12), k14: n0(r.a14) + n0(r.s14),
    k18: n0(r.a18) + n0(r.s18), k22: n0(r.a22) + n0(r.s22),
  };
  const defKg = DEFECTS.reduce((s, [k]) => s + n0(r[k]), 0);
  const groups = {};
  for (const [g, keys] of Object.entries(GROUPS))
    groups[g] = p ? keys.reduce((s, k) => s + n0(r[k]), 0) / p : null;
  return {
    air, sea, ...sz,
    packed: air + sea,
    returnPct: reaped && p != null ? (p - reaped) / reaped : null,
    airPct: p ? air / p : null,
    seaPct: p ? sea / p : null,
    p12: p ? sz.k12 / p : null, p14: p ? sz.k14 / p : null,
    p18: p ? sz.k18 / p : null, p22: p ? sz.k22 / p : null,
    pu12: p ? n0(r.under12) / p : null,
    defKg, defPct: p ? defKg / p : null, groups,
    complete: p != null && r.variety && r.block,
  };
}

const kg = (v, d = 0) =>
  v == null ? "—" : v.toLocaleString("en", { minimumFractionDigits: d, maximumFractionDigits: d });
const pc = (v, d = 1) => (v == null ? "—" : (v * 100).toFixed(d) + "%");
const dfmt = (s) => {
  if (!s) return "—";
  const d = new Date(s + "T00:00:00");
  return d.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
};
const dlong = (s) => {
  if (!s) return "—";
  const d = new Date(s + "T00:00:00");
  return d.toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short", year: "numeric" });
};
function isoWeek(s) {
  const d = new Date(s + "T00:00:00");
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t - y0) / 86400000 + 1) / 7);
}

const SIZE_COLORS = { u12: "#D9DBF0", k12: "#B9BEEC", k14: "#7F86DE", k18: "#4C46A8", k22: "#28244E" };
const GROUP_COLORS = { "Reaping issues": "#AE4038", "Under ripe": "#B4761F", "Over ripe": "#7A4E9E", "Other": "#5B7A99" };

/* ================= PACKSHED UPDATE (from Packshed_Update_12.xlsx) ================= */
const PACKSHED = JSON.parse('{"asOf":"2026-07-15T06:57","week":29,"issue":1,"orders":[{"customer":"DRIS SEA 1","size":"12MM+","unit":3.5,"ord":[0,4368,6650],"packed":[0,4368,3188.5],"bal":[0,0,3461.5]},{"customer":"DRIS SEA 2","size":"14MM+","unit":3.5,"ord":[9800,0,0],"packed":[9800,0,0],"bal":[0,0,0]},{"customer":"DRIS SEA 3","size":"18MM+","unit":3.5,"ord":[0,1960,2135],"packed":[0,1960,2135],"bal":[0,0,0]},{"customer":"EGR","size":"18MM+","unit":2,"ord":[2520,0,0],"packed":[2520,0,0],"bal":[0,0,0]},{"customer":"FREC-D","size":"18MM+","unit":2,"ord":[840,0,0],"packed":[840,0,0],"bal":[0,0,0]},{"customer":"FREC-D","size":"22MM+","unit":1.5,"ord":[504,0,0],"packed":[504,0,0],"bal":[0,0,0]},{"customer":"GFP SEA","size":"12MM+","unit":3.5,"ord":[1960,1522.5,0],"packed":[1960,1522.5,0],"bal":[0,0,0]},{"customer":"HIH","size":"18MM+","unit":2,"ord":[420,840,0],"packed":[420,840,0],"bal":[0,0,0]},{"customer":"HIH","size":"22MM+","unit":1.5,"ord":[57,309,0],"packed":[57,309,0],"bal":[0,0,0]},{"customer":"NTB","size":"18MM+","unit":2,"ord":[1680,0,0],"packed":[1680,0,0],"bal":[0,0,0]},{"customer":"RICH-D","size":"18MM+","unit":2,"ord":[0,1260,0],"packed":[0,1260,0],"bal":[0,0,0]},{"customer":"RICH-D","size":"22MM+","unit":1.5,"ord":[0,0,0],"packed":[0,0,0],"bal":[0,0,0]},{"customer":"EURO","size":"ALL","unit":9,"ord":[0,0,2520],"packed":[0,0,1170],"bal":[0,0,1350]},{"customer":"CHINA","size":"18MM+","unit":1.5,"ord":[108,0,0],"packed":[108,0,0],"bal":[0,0,0]}],"varieties":[{"name":"AKALA","g":[0,0.6,0.3,0,0.1],"stock":1201,"proj":[120.1,0,720.6,360.3,0]},{"name":"ALLIRA","g":[0,0.16,0.74,0,0.1],"stock":0,"proj":[0,0,0,0,0]},{"name":"ARANA","g":[0,0.32,0.4,0.2,0.08],"stock":0,"proj":[0,0,0,0,0]},{"name":"BOUNTY","g":[0,0.1,0.65,0.2,0.05],"stock":0,"proj":[0,0,0,0,0]},{"name":"BREEZE","g":[0,0.12,0.8,0,0.08],"stock":0,"proj":[0,0,0,0,0]},{"name":"CASCADE","g":[0,0.25,0.55,0.1,0.1],"stock":0,"proj":[0,0,0,0,0]},{"name":"DELIGHT","g":[0,0.5,0.4,0,0.1],"stock":0,"proj":[0,0,0,0,0]},{"name":"ETERNA","g":[0,0.2,0.65,0.05,0.1],"stock":0,"proj":[0,0,0,0,0]},{"name":"KIRRA OPEN","g":[0.06,0.74,0.08,0,0.12],"stock":833,"proj":[99.96,49.98,616.42,66.64,0]},{"name":"KIRRA PLASTIC","g":[0,0.63,0.25,0,0.12],"stock":0,"proj":[0,0,0,0,0]},{"name":"PROMISE","g":[0,0.9,0,0,0.1],"stock":0,"proj":[0,0,0,0,0]},{"name":"VERDURE","g":[0.06,0.69,0.15,0,0.1],"stock":0,"proj":[0,0,0,0,0]}],"unprocTotal":[220.06,49.98,1337.02,426.94,0,2034],"procNotPacked":[0,0,0,381,0,381],"packedBuffer":[0,0,0,2034,2457,4491],"orderedBySize":{"sat":[0,1960,9800,5568,561,17889],"tue":[0,5890.5,0,4060,309,10259.5],"thu":[2520,6650,0,2135,0,11305],"total":[2520,14500.5,9800,11763,870,39453.5]},"packedBySize":{"total":[1170,11039,9800,11763,870,34642]},"balanceBySize":{"total":[1350,3461.5,0,0,0,4811.5]},"available":[220.06,49.98,1337.02,2841.94,2457,6906],"surplus":[-1129.94,-3411.52,1337.02,2841.94,2457,2094.5],"localSales":[{"item":"Local sales 5 kg crate","cartons":586,"kgs":2930},{"item":"Local sales 2 kg box","cartons":12,"kgs":64}]}');
const PK_SIZES = ["ALL", "12mm+", "14mm+", "18mm+", "22mm+"];
// Preset grade-out banding per variety (from the packshed workbook), plus any
// varieties seen in batch data, for the stock-line variety dropdown.
const VARIETY_PRESETS = Object.fromEntries(PACKSHED.varieties.map((v) => [v.name, v.g]));
const STOCK_VARIETY_OPTS = [...new Set([
  ...PACKSHED.varieties.map((v) => v.name),
  ...SEED.rows.map((r) => r[idx.variety]).filter((v) => v && v !== "ALL" && v !== "REJECTS"),
])].sort();

/* ================= STORAGE ================= */
const BATCH_KEY = "ml-batches-shared-v1";      // shared: one team-wide batch register
const LEGACY_BATCH_KEY = "ml-user-batches-v1"; // old per-account store, migrated on load
async function loadUser() {
  let shared = [];
  try {
    const r = await window.storage.get(BATCH_KEY, true);
    shared = r ? JSON.parse(r.value) : [];
  } catch {}
  // one-time migration: move this account's old personal batches into the shared register
  try {
    const r = await window.storage.get(LEGACY_BATCH_KEY);
    const legacy = r ? JSON.parse(r.value) : [];
    if (legacy.length) {
      const ids = new Set(shared.map((x) => x.id));
      const add = legacy.filter((x) => !ids.has(x.id));
      if (add.length) {
        shared = [...shared, ...add];
        await window.storage.set(BATCH_KEY, JSON.stringify(shared), true);
      }
      await window.storage.delete(LEGACY_BATCH_KEY);
    }
  } catch {}
  return shared;
}
async function saveUser(list) {
  try { await window.storage.set(BATCH_KEY, JSON.stringify(list), true); return true; }
  catch { return false; }
}

/* ================= AUTH ================= */
const USERS_KEY = "ml-users-v1";     // shared: one team-wide user list
const SESSION_KEY = "ml-session-v1"; // personal: this device's sign-in
const ROLE_LABEL = { admin: "Admin", entry: "Data entry", viewer: "View only", qc: "QC", hrOfficer: "HR Officer", hrManager: "HR Manager" };

// Pure-JS SHA-256 (verified identical to crypto.subtle output) — fallback for
// embedded contexts where crypto.subtle is unavailable, so sign-in never hangs.
function shaJS(str) {
  const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
  let H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
  const bytes = new TextEncoder().encode(str);
  const l = bytes.length, bitLen = l * 8;
  const padded = new Uint8Array((((l + 8) >> 6) << 6) + 64);
  padded.set(bytes); padded[l] = 0x80;
  const dv = new DataView(padded.buffer);
  dv.setUint32(padded.length - 4, bitLen >>> 0);
  dv.setUint32(padded.length - 8, Math.floor(bitLen / 4294967296));
  const w = new Uint32Array(64);
  const rr = (x, n) => (x >>> n) | (x << (32 - n));
  for (let off = 0; off < padded.length; off += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getUint32(off + i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rr(w[i - 15], 7) ^ rr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rr(w[i - 2], 17) ^ rr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }
    let [a, b, c, d, e, f, g, h] = H;
    for (let i = 0; i < 64; i++) {
      const S1 = rr(e, 6) ^ rr(e, 11) ^ rr(e, 25);
      const ch = (e & f) ^ (~e & g);
      const t1 = (h + S1 + ch + K[i] + w[i]) >>> 0;
      const S0 = rr(a, 2) ^ rr(a, 13) ^ rr(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (S0 + maj) >>> 0;
      h = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = b; b = a; a = (t1 + t2) >>> 0;
    }
    const out = [a, b, c, d, e, f, g, h];
    H = H.map((x, i) => (x + out[i]) >>> 0);
  }
  return H.map((x) => x.toString(16).padStart(8, "0")).join("");
}
async function sha(text) {
  const msg = "ml·" + text;
  try {
    if (typeof crypto !== "undefined" && crypto.subtle && crypto.subtle.digest) {
      const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(msg));
      return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
    }
  } catch {}
  return shaJS(msg);
}
async function loadUsersStore() {
  try { const r = await window.storage.get(USERS_KEY, true); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function saveUsersStore(list) {
  try { await window.storage.set(USERS_KEY, JSON.stringify(list), true); return true; }
  catch { return false; }
}
async function loadSession() {
  try { const r = await window.storage.get(SESSION_KEY); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function saveSession(sess) { try { await window.storage.set(SESSION_KEY, JSON.stringify(sess)); } catch {} }
async function clearSession() { try { await window.storage.delete(SESSION_KEY); } catch {} }

/* ================= SMALL UI ================= */
const Kpi = ({ label, value, sub, tone }) => (
  <div className={"kpi" + (tone ? " kpi-" + tone : "")}>
    <div className="kpi-label">{label}</div>
    <div className="kpi-value">{value}</div>
    {sub != null && <div className="kpi-sub">{sub}</div>}
  </div>
);

const Card = ({ title, right, children, pad = true }) => (
  <section className="card">
    {(title || right) && (
      <header className="card-head">
        <h3>{title}</h3>
        <div>{right}</div>
      </header>
    )}
    <div className={pad ? "card-body" : ""}>{children}</div>
  </section>
);

const SNAP_LIBS = [
  ["html2canvas", [
    "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js",
    "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js",
    "https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js"]],
  ["jspdf", [
    "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
    "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js",
    "https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js"]],
];
async function loadSnapLibs() {
  for (const [g, urls] of SNAP_LIBS) {
    if (window[g]) continue;
    const code = await pdfFetchFirst(urls);
    const sc = document.createElement("script");
    sc.textContent = code;
    document.head.appendChild(sc);
    if (!window[g]) throw new Error("pdf-lib-net");
  }
  return true;
}

function PdfBtn({ name }) {
  const [busy, setBusy] = useState(false);
  const [note, setNote] = useState("");
  const click = async () => {
    if (busy) return;
    setBusy(true); setNote("");
    try {
      await loadSnapLibs(); // first use downloads two small components, then cached
      document.body.classList.add("snap");
      await new Promise((r) => setTimeout(r, 80)); // let the letterhead paint
      const el = document.querySelector(".main");
      const capW = Math.max(el.scrollWidth, 1000);
      const canvas = await window.html2canvas(el, {
        scale: 2, backgroundColor: "#ffffff", useCORS: true, width: capW, windowWidth: capW,
      });
      document.body.classList.remove("snap");
      const { jsPDF } = window.jspdf;
      // wide content (e.g. the Quality table) fits better sideways
      const landscape = canvas.width > canvas.height;
      const pdf = new jsPDF({ orientation: landscape ? "landscape" : "portrait", unit: "mm", format: "a4" });
      const M = 8, availW = (landscape ? 297 : 210) - 2 * M, availH = (landscape ? 210 : 297) - 2 * M;
      const ratio = canvas.height / canvas.width;
      let w = availW, h = availW * ratio;
      if (h > availH) { h = availH; w = availH / ratio; } // always exactly one page
      pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", M + (availW - w) / 2, M, w, h);
      pdf.save("Merrylight_" + (name || "Dashboard") + "_" + new Date().toISOString().slice(0, 10) + ".pdf");
    } catch (e) {
      document.body.classList.remove("snap");
      setNote(String((e && e.message) === "pdf-lib-net"
        ? "Couldn't download the PDF component \u2014 first use needs internet; check the connection and retry."
        : "PDF failed: " + ((e && e.message) || e) + " \u2014 screenshot this and send it to Claude."));
    }
    setBusy(false);
  };
  return (
    <>
      <button className="btn btn-mini" style={{ padding: "7px 12px" }} disabled={busy} onClick={click}>
        {busy ? "Preparing\u2026" : "Download PDF"}
      </button>
      {note && <span className="savedmsg warn" style={{ marginTop: 0, maxWidth: 280, whiteSpace: "normal" }}>{note}</span>}
    </>
  );
}

const Select = ({ value, onChange, options, all }) => (
  <select className="sel" value={value} onChange={(e) => onChange(e.target.value)}>
    {all && <option value="">{all}</option>}
    {options.map((o) => <option key={o} value={o}>{o}</option>)}
  </select>
);


/* ================= DASHBOARD ================= */
function Dashboard({ recs }) {
  const dates = useMemo(() => {
    const s = [...new Set(recs.filter((r) => num(r.processed)).map((r) => r.date))];
    return s.sort().reverse(); // newest first
  }, [recs]);
  const [date, setDate] = useState(dates[0] || "");
  useEffect(() => { if (!dates.includes(date)) setDate(dates[0] || ""); }, [dates]);

  const day = recs.filter((r) => r.date === date && num(r.processed) != null);
  const dayD = day.map(derive);
  const tot = day.reduce((s, r) => s + n0(r.processed), 0);
  const reap = day.reduce((s, r) => s + n0(r.reaped), 0);
  const air = dayD.reduce((s, d) => s + d.air, 0);
  const sea = dayD.reduce((s, d) => s + d.sea, 0);
  const sz = ["k12", "k14", "k18", "k22"].map((k) => dayD.reduce((s, d) => s + d[k], 0));
  const u12 = day.reduce((s, r) => s + n0(r.under12), 0);
  const packed = air + sea + u12;
  const bw = (() => {
    const w = day.filter((r) => num(r.bw) && num(r.processed));
    const t = w.reduce((s, r) => s + r.processed, 0);
    return t ? w.reduce((s, r) => s + r.bw * r.processed, 0) / t : null;
  })();

  // size grade by block (share of processed kg), for the selected date
  const byBlock = {};
  day.forEach((r) => {
    const d = derive(r);
    const b = (byBlock[r.block] ||= { block: r.block, p: 0, k12: 0, k14: 0, k18: 0, k22: 0 });
    b.p += n0(r.processed); b.k12 += d.k12; b.k14 += d.k14; b.k18 += d.k18; b.k22 += d.k22;
  });
  const blockChart = Object.values(byBlock).filter((b) => b.p > 0).map((b) => ({
    block: b.block,
    "12mm+": +((b.k12 / b.p) * 100).toFixed(1),
    "14mm+": +((b.k14 / b.p) * 100).toFixed(1),
    "18mm+": +((b.k18 / b.p) * 100).toFixed(1),
    "22mm+": +((b.k22 / b.p) * 100).toFixed(1),
  }));

  // processed kg by variety for the selected date ("ALL"/"REJECTS" lines excluded to avoid double counting)
  const byVar = {};
  day.forEach((r) => {
    if (!r.variety || r.variety === "ALL" || r.variety === "REJECTS") return;
    byVar[r.variety] = (byVar[r.variety] || 0) + n0(r.processed);
  });
  const varChart = Object.entries(byVar)
    .map(([variety, kgs]) => ({ variety, kgs: +kgs.toFixed(0) }))
    .sort((a, b) => a.variety.localeCompare(b.variety));

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Daily dashboard</h2>
          <p className="muted">{dlong(date)} · Week {date ? isoWeek(date) : "—"} · {day.length} batch{day.length === 1 ? "" : "es"}</p>
        </div>
        <div className="controls">
          <label className="lbl">Date</label>
          <Select value={date} onChange={setDate} options={dates} />
          <PdfBtn name="Daily_Dashboard" />
        </div>
      </div>

      <div className="kpis">
        <Kpi label="Processed" value={kg(tot) + " kg"} sub={reap ? "Reaped " + kg(reap) + " kg" : " "} />
        <Kpi label="Return" value={reap ? pc((tot - reap) / reap) : "—"} tone={reap && tot < reap ? "bad" : "good"} sub="(processed − reaped) / reaped" />
        <Kpi label="Air / Sea" value={packed ? pc(air / (air + sea || 1), 0) + " / " + pc(sea / (air + sea || 1), 0) : "—"} sub={kg(air) + " kg air · " + kg(sea) + " kg sea"} />
        <Kpi label="Avg berry weight" value={bw ? bw.toFixed(2) + " g" : "—"} sub="weighted by batch kg" />
        <Kpi label="Blocks / Varieties" value={new Set(day.map((r) => r.block)).size + " / " + new Set(day.map((r) => r.variety)).size} sub={" "} />
      </div>

      <div className="grid2">
        <Card title="Size grade by block" right={<span className="muted small">% of processed kg</span>}>
          {blockChart.length ? (
            <SBarChart data={blockChart} xKey="block" height={280} unit="%"
              series={[
                { key: "12mm+", color: "#4472C4" },
                { key: "14mm+", color: "#ED7D31" },
                { key: "18mm+", color: "#A5A5A5" },
                { key: "22mm+", color: "#FFC000" },
              ]} />
          ) : <p className="muted">No grading data for this date.</p>}
        </Card>

        <Card title="Processed kgs by variety" right={<span className="muted small">{dfmt(date)} · kg</span>}>
          {varChart.length ? (
            <SBarChart data={varChart} xKey="variety" height={280} fmt={(v) => kg(v)}
              series={[{ key: "kgs", label: "Processed", color: "#4472C4" }]} />
          ) : <p className="muted">No processed kg recorded for this date.</p>}
        </Card>
      </div>

      <Card title="Batches on this date" pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>Batch</th><th>Block</th><th>Variety</th>
                <th className="r">Processed kg</th>
                <th className="r">Berry wt</th><th className="r">12mm+</th><th className="r">14mm+</th><th className="r">18mm+</th><th className="r">22mm+</th>
                <th className="r">Air</th><th className="r">Sea</th><th className="r">Defects</th>
              </tr>
            </thead>
            <tbody>
              {day.map((r) => {
                const d = derive(r);
                return (
                  <tr key={r.id} className={d.defPct > 0.1 ? "hidef" : ""}>
                    <td className="mono">{r.batch || "—"}</td>
                    <td>{r.block}</td><td>{r.variety}</td>
                    <td className="r strong">{kg(r.processed, 1)}</td>
                    
                    <td className="r">{r.bw ?? "—"}</td>
                    <td className="r">{pc(d.p12)}</td><td className="r">{pc(d.p14)}</td><td className="r">{pc(d.p18)}</td><td className="r">{pc(d.p22)}</td>
                    <td className="r">{pc(d.airPct)}</td>
                    <td className="r">{pc(d.seaPct)}</td>
                    <td className={"r" + (d.defPct > 0.1 ? " neg strong" : "")}>{pc(d.defPct)}</td>
                  </tr>
                );
              })}
              {!day.length && <tr><td colSpan={12} className="muted center">No batches recorded on this date.</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </>
  );
}

/* ================= BATCHES ================= */
function Batches({ recs, onDelete, me }) {
  const [q, setQ] = useState("");
  const [variety, setVariety] = useState("");
  const [block, setBlock] = useState("");
  const [week, setWeek] = useState("");
  const [open, setOpen] = useState(null);

  const varieties = [...new Set(recs.map((r) => r.variety).filter(Boolean))].sort();
  const blocks = [...new Set(recs.map((r) => r.block).filter(Boolean))].sort();
  const weeks = [...new Set(recs.map((r) => r.date && isoWeek(r.date)).filter(Boolean))].sort((a, b) => b - a); // newest first

  const list = recs
    .filter((r) =>
      (!q || (r.batch || "").toLowerCase().includes(q.toLowerCase()) || (r.block || "").toLowerCase().includes(q.toLowerCase())) &&
      (!variety || r.variety === variety) && (!block || r.block === block) &&
      (!week || (r.date && isoWeek(r.date) === +week)))
    .sort((a, b) => (b.date || "").localeCompare(a.date || "") || (b.batch || "").localeCompare(a.batch || ""));

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Batch records</h2>
          <p className="muted">{list.length} of {recs.length} records</p>
        </div>
        <div className="controls wrap">
          <input className="inp" placeholder="Search batch or block…" value={q} onChange={(e) => setQ(e.target.value)} />
          <Select value={variety} onChange={setVariety} options={varieties} all="All varieties" />
          <Select value={block} onChange={setBlock} options={blocks} all="All blocks" />
          <Select value={week} onChange={setWeek} options={weeks.map(String)} all="All weeks" />
        </div>
      </div>

      <Card pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>Date</th><th>Batch</th><th>Block</th><th>Variety</th>
                <th className="r">Processed kg</th>
                <th className="r">12mm+</th><th className="r">14mm+</th><th className="r">18mm+</th><th className="r">22mm+</th>
                <th className="r">Defects</th><th></th>
              </tr>
            </thead>
            <tbody>
              {list.map((r) => {
                const d = derive(r);
                const isOpen = open === r.id;
                return (
                  <React.Fragment key={r.id}>
                    <tr className={"rowbtn" + (r.source === "app" ? " userrow" : "") + (d.defPct > 0.1 ? " hidef" : "")} onClick={() => setOpen(isOpen ? null : r.id)}>
                      <td>{dfmt(r.date)}</td>
                      <td className="mono">{r.batch || "—"}{r.source === "app" && <span className="badge">app</span>}</td>
                      <td>{r.block || "—"}</td><td>{r.variety || "—"}</td>
                      <td className="r strong">{kg(r.processed, 1)}</td>
                      
                      <td className="r">{pc(d.p12)}</td><td className="r">{pc(d.p14)}</td><td className="r">{pc(d.p18)}</td><td className="r">{pc(d.p22)}</td>
                      <td className={"r" + (d.defPct > 0.1 ? " neg strong" : "")}>{pc(d.defPct)}</td>
                      <td className="r chev">{isOpen ? "▾" : "▸"}</td>
                    </tr>
                    {isOpen && (
                      <tr className="detail">
                        <td colSpan={11}>
                          <div className="detail-grid">
                            <div>
                              <h4>Record</h4>
                              <dl>
                                <dt>Date</dt><dd>{dlong(r.date)}</dd>
                                <dt>Shift / Week</dt><dd>{r.shift || "—"} · Wk {r.date ? isoWeek(r.date) : "—"}</dd>
                                <dt>Section</dt><dd>{r.section || "—"}</dd>
                                <dt>Reaped</dt><dd>{kg(r.reaped, 1)} kg</dd>
                                <dt>Processed</dt><dd>{kg(r.processed, 1)} kg</dd>
                                <dt>Berry wt / kg·h</dt><dd>{r.bw ?? "—"} g · {kg(r.kgh)}</dd>
                                {r.enteredBy && <><dt>Entered by</dt><dd>{r.enteredBy}</dd></>}
                              </dl>
                            </div>
                            <div>
                              <h4>Grading (kg)</h4>
                              <dl>
                                <dt>12+ air / sea</dt><dd>{kg(r.a12, 1)} / {kg(r.s12, 1)}</dd>
                                <dt>14+ air / sea</dt><dd>{kg(r.a14, 1)} / {kg(r.s14, 1)}</dd>
                                <dt>18+ air / sea</dt><dd>{kg(r.a18, 1)} / {kg(r.s18, 1)}</dd>
                                <dt>22+ air / sea</dt><dd>{kg(r.a22, 1)} / {kg(r.s22, 1)}</dd>
                                <dt>Air / Sea total</dt><dd>{kg(d.air, 1)} / {kg(d.sea, 1)}</dd>
                              </dl>
                            </div>
                            <div>
                              <h4>Defects (kg)</h4>
                              <dl>
                                {DEFECTS.filter(([k]) => n0(r[k]) > 0).map(([k, l]) => (
                                  <React.Fragment key={k}><dt>{l}</dt><dd>{kg(r[k], 1)}</dd></React.Fragment>
                                ))}
                                {d.defKg === 0 && <><dt>Total</dt><dd>0</dd></>}
                              </dl>
                            </div>
                            {r.source === "app" && me.role === "admin" && (
                              <div className="detail-actions">
                                <button className="btn btn-danger" onClick={(e) => { e.stopPropagation(); onDelete(r.id); }}>
                                  Delete record
                                </button>
                              </div>
                            )}
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
              {!list.length && <tr><td colSpan={11} className="muted center">No records match these filters.</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </>
  );
}

/* ================= NEW BATCH ================= */
const EMPTY = {
  date: new Date().toISOString().slice(0, 10), shift: "DAY", batch: "", block: "", variety: "",
  section: "COMPLEX", reaped: "", processed: "", bw: "", kgh: "",
  a12: "", a14: "", a18: "", a22: "", s12: "", s14: "", s18: "", s22: "",
  tears: "", redBums: "", greenBums: "", mold: "", stemPull: "", shrivell: "", scar: "",
  softPoint: "", soft: "", under12: "", subBloom: "",
};

/* ================= FILE IMPORT ================= */
const IMPORT_COLS = {
  "Date Processed": "date", "Shift": "shift", "Batch Number": "batch", "Block": "block",
  "Variety": "variety", "Section": "section", "Reaped Kgs": "reaped", "Processed Kgs": "processed",
  "Av. Berry Weight": "bw", "Kgs / Hour": "kgh",
  "12+ Air": "a12", "14+ Air": "a14", "18+ Air": "a18", "22+ Air": "a22",
  "12+ Sea": "s12", "14+ Sea": "s14", "18+ Sea": "s18", "22+ Sea": "s22",
  "Tears": "tears", "Red Bums": "redBums", "Green Bums": "greenBums", "Mold": "mold",
  "Stem Pull": "stemPull", "Shrivell": "shrivell", "Scar": "scar", "Soft Point": "softPoint",
  "Soft": "soft", "<12mm+": "under12", "Sub Bloom": "subBloom",
};
const TEXT_FIELDS = ["date", "shift", "batch", "block", "variety", "section"];

function toISODate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return isNaN(v) ? null : v.toISOString().slice(0, 10);
  if (typeof v === "number" && v > 20000 && v < 60000) { // Excel serial date
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return isNaN(d) ? null : d.toISOString().slice(0, 10);
  }
  const str = String(v).trim();
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  const MON = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  const mm = str.match(/^(?:[A-Za-z]{3,9}[ ,]+)?(\d{1,2})[-\/ ]([A-Za-z]{3,9})[-\/ ](\d{2,4})$/);
  if (mm && MON[mm[2].slice(0, 3).toLowerCase()] != null) {
    const y = mm[3].length === 2 ? 2000 + +mm[3] : +mm[3];
    const d2 = new Date(Date.UTC(y, MON[mm[2].slice(0, 3).toLowerCase()], +mm[1]));
    return isNaN(d2) ? null : d2.toISOString().slice(0, 10);
  }
  const dmy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // day-first (en-GB)
  if (dmy) {
    const d3 = new Date(Date.UTC(+dmy[3], +dmy[2] - 1, +dmy[1]));
    return isNaN(d3) ? null : d3.toISOString().slice(0, 10);
  }
  const d = new Date(str);
  return isNaN(d) ? null : d.toISOString().slice(0, 10);
}

function parseBatchFile(data) {
  const wb = XLSX.read(data, { type: "array" });
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    const hi = rows.findIndex((r) => r && r.includes("Date Processed") && r.includes("Batch Number"));
    if (hi === -1) continue;
    const colIdx = {};
    rows[hi].forEach((h, i) => {
      const key = typeof h === "string" ? h.trim() : h;
      if (IMPORT_COLS[key] && colIdx[IMPORT_COLS[key]] == null) colIdx[IMPORT_COLS[key]] = i;
    });
    const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const get = (f) => (colIdx[f] != null ? r[colIdx[f]] : null);
      const date = toISODate(get("date"));
      if (!date) continue;
      const rec = { date, week: isoWeek(date) };
      for (const f of Object.values(IMPORT_COLS)) {
        if (f === "date") continue;
        const v = get(f);
        rec[f] = TEXT_FIELDS.includes(f)
          ? (v == null || v === "" ? null : String(v).trim().toUpperCase())
          : num(v);
      }
      if (!rec.batch && rec.processed == null) continue; // empty template row
      out.push(rec);
    }
    if (out.length) return { sheet: name, rows: out };
  }
  return null;
}

/* ================= LOCAL PDF EXTRACTION (pdf.js) =================
   PDFs are read entirely on the device: no API key, no per-file cost.
   Works on text-based PDFs (e.g. exported/printed from Excel).
   Scanned paper or photos have no text layer and cannot be read. */
const PDFJS_SRC = {
  lib: [
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
    "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js",
    "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js",
  ],
  worker: [
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",
    "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js",
    "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js",
  ],
};
async function pdfFetchFirst(urls) {
  for (const u of urls) {
    try {
      const r = await fetch(u);
      if (r.ok) return await r.text();
    } catch (e) { /* next mirror */ }
  }
  throw new Error("pdf-lib-net");
}
function pdfInject(code) {
  const sc = document.createElement("script");
  sc.textContent = code;
  document.head.appendChild(sc);
}
async function loadPdfJs() {
  if (typeof window !== "undefined" && window.pdfjsLib && window.pdfjsWorker) return window.pdfjsLib;
  if (!window.pdfjsLib) {
    pdfInject(await pdfFetchFirst(PDFJS_SRC.lib));
    if (!window.pdfjsLib) throw new Error("pdf-lib-net");
  }
  // No blob/URL worker loading (blocked in sandboxed environments):
  // evaluate the worker inline and register it as the main-thread handler.
  if (!window.pdfjsWorker || !window.pdfjsWorker.WorkerMessageHandler) {
    pdfInject(await pdfFetchFirst(PDFJS_SRC.worker));
    window.pdfjsWorker = window.pdfjsWorker || window["pdfjs-dist/build/pdf.worker"];
    if (!window.pdfjsWorker || !window.pdfjsWorker.WorkerMessageHandler) throw new Error("pdf-lib-net");
  }
  return window.pdfjsLib;
}

// --- pure table reconstruction (verifiable outside the browser) ---
const pdfNorm = (t) => String(t).toLowerCase().replace(/[^a-z0-9<+]/g, "");
function pdfLines(items) {
  // cluster positioned text items into visual lines, merging fragments of one cell
  const lines = [];
  for (const it of items) {
    if (!it.str || !it.str.trim()) continue;
    let line = lines.find((l) => Math.abs(l.y - it.y) < 2.5);
    if (!line) { line = { y: it.y, cells: [] }; lines.push(line); }
    line.cells.push({ x: it.x, xEnd: it.x + (it.w || it.str.length * 4), str: it.str.trim() });
  }
  lines.sort((a, b) => b.y - a.y); // PDF y grows upward: top of page first
  for (const l of lines) {
    l.cells.sort((a, b) => a.x - b.x);
    const merged = [];
    for (const c of l.cells) {
      const last = merged[merged.length - 1];
      if (last && c.x - last.xEnd < 6) { last.str += " " + c.str; last.xEnd = c.xEnd; }
      else merged.push({ ...c });
    }
    l.cells = merged;
  }
  return lines;
}
function pdfRecords(lines) {
  const hi = lines.findIndex((l) => {
    const t = pdfNorm(l.cells.map((c) => c.str).join(" "));
    return t.includes("dateprocessed") && t.includes("batchnumber");
  });
  if (hi === -1) return null;
  const anchors = lines[hi].cells.map((c) => ({ x: (c.x + c.xEnd) / 2, text: c.str }));

  // absorb wrapped header lines (tall header cells print as extra lines) until real data starts
  let di = hi + 1, absorbed = 0;
  while (di < lines.length && absorbed < 3) {
    const l = lines[di];
    if (l.cells.length && toISODate(l.cells[0].str)) break; // data rows start with a date
    for (const c of l.cells) {
      let best = null, bd = 1e9;
      for (const a of anchors) {
        const d = Math.abs((c.x + c.xEnd) / 2 - a.x);
        if (d < bd) { bd = d; best = a; }
      }
      if (best && bd < 60) best.text += " " + c.str;
    }
    di++; absorbed++;
  }

  // map anchors to import fields: exact match first, then contains
  const colFor = {};
  const entries = Object.entries(IMPORT_COLS);
  for (const pass of [0, 1]) {
    anchors.forEach((a, ai) => {
      const nt = pdfNorm(a.text);
      for (const [h, f] of entries) {
        if (colFor[f] != null) continue;
        const nh = pdfNorm(h);
        if (pass === 0 ? nt === nh : nt.includes(nh)) { colFor[f] = ai; break; }
      }
    });
  }
  if (colFor.date == null || colFor.batch == null) return null;

  const out = [];
  for (let i = di; i < lines.length; i++) {
    const assign = {};
    for (const c of lines[i].cells) {
      let bi = -1, bd = 1e9;
      anchors.forEach((a, ai) => {
        const d = Math.abs((c.x + c.xEnd) / 2 - a.x);
        if (d < bd) { bd = d; bi = ai; }
      });
      if (bi >= 0) assign[bi] = assign[bi] ? assign[bi] + " " + c.str : c.str;
    }
    const get = (f) => (colFor[f] != null && assign[colFor[f]] != null ? assign[colFor[f]] : null);
    const date = toISODate(get("date"));
    if (!date) continue; // totals, repeated headers, footers
    const rec = { date, week: isoWeek(date) };
    for (const f of Object.values(IMPORT_COLS)) {
      if (f === "date") continue;
      const v = get(f);
      rec[f] = TEXT_FIELDS.includes(f)
        ? (v == null || v === "" ? null : String(v).trim().toUpperCase())
        : num(String(v == null ? "" : v).replace(/,/g, ""));
    }
    if (!rec.batch && rec.processed == null) continue;
    out.push(rec);
  }
  return out.length ? { sheet: "PDF", rows: out } : null;
}

/* --- Profile 2: TrueSort / Ellips Classification Report (one batch per PDF) ---
   Text layer quirks handled: "+" signs, hyphens, and colons are dropped by the
   generator ("14+ Sea" -> "14 Sea", "MER00361-C04" -> "MER00361C04", "7:08" -> "708"). */
const GRADE_MAP = {
  "12sea": "s12", "14sea": "s14", "18sea": "s18", "22sea": "s22",
  "12air": "a12", "14air": "a14", "18air": "a18", "22air": "a22",
  tear: "tears", tears: "tears", red: "redBums", redbum: "redBums", redbums: "redBums",
  green: "greenBums", greenbum: "greenBums", greenbums: "greenBums",
  mold: "mold", mould: "mold", stem: "stemPull", stempull: "stemPull",
  shriveled: "shrivell", shrivelled: "shrivell", shrivel: "shrivell",
  scar: "scar", softpoint: "softPoint", soft: "soft",
  "1012": "under12", "<12mm": "under12", subbloom: "subBloom",
};
const gnorm = (t) => String(t).toLowerCase().replace(/[^a-z0-9<]/g, "");

function pdfClassification(lines) {
  const KEYS = {
    batch: ["batch", /\d/], variety: ["variety", null], block: ["block", null],
    processed: ["totalweightkg", /\d/], processed2: ["totalweight", /\d/],
    start: ["starttime", /\d/], end: ["endtime", /\d/],
  };
  const kv = {};
  for (const l of lines) {
    for (let i = 0; i < l.cells.length - 1; i++) {
      const kn = gnorm(l.cells[i].str);
      for (const [field, [key, valid]] of Object.entries(KEYS)) {
        const tgt = field === "processed2" ? "processed" : field;
        if (kn === key && kv[tgt] == null) {
          const v = l.cells[i + 1].str;
          if (!valid || valid.test(v)) kv[tgt] = v;
        }
      }
    }
  }
  if (kv.batch == null || kv.processed == null || kv.start == null) return null;

  const dm = String(kv.start).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); // machine uses M/D/YYYY
  if (!dm) return null;
  let mo = +dm[1], dy = +dm[2];
  if (mo > 12 && dy <= 12) { const t = mo; mo = dy; dy = t; }
  const date = new Date(Date.UTC(+dm[3], mo - 1, dy)).toISOString().slice(0, 10);

  const clock = (s2) => { // "7:08 AM" or colon-stripped "708 AM"
    const m2 = String(s2).match(/(\d{1,2})[:.]?(\d{2})\s*(AM|PM)/i);
    if (!m2) return null;
    let h = +m2[1] % 12;
    if (/pm/i.test(m2[3])) h += 12;
    return h + +m2[2] / 60;
  };
  const st = clock(kv.start), en = clock(kv.end);
  const processed = num(String(kv.processed).replace(/,/g, ""));
  const hours = st != null && en != null && en > st ? en - st : null;

  const rawB = String(kv.batch).toUpperCase().replace(/[^A-Z0-9]/g, "");
  const blk = kv.block ? String(kv.block).toUpperCase().replace(/[^A-Z0-9]/g, "") : null;
  const batch = blk && rawB.endsWith(blk) && rawB.length > blk.length
    ? rawB.slice(0, rawB.length - blk.length) : rawB;

  const rec = {
    date, week: isoWeek(date),
    shift: st != null ? (st >= 4 && st < 16 ? "DAY" : "NIGHT") : null,
    batch, block: blk,
    variety: kv.variety ? String(kv.variety).trim().toUpperCase() : null,
    section: null, reaped: null, processed, bw: null,
    kgh: hours && processed != null ? +(processed / hours).toFixed(1) : null,
  };
  for (const f of Object.values(IMPORT_COLS)) if (rec[f] === undefined) rec[f] = null;

  // GRADE table: header -> its Total row; weights land on their register fields
  const hi = lines.findIndex((l) => {
    const t = gnorm(l.cells.map((c) => c.str).join(""));
    return t.includes("grade") && t.includes("count") && t.includes("totalweight");
  });
  if (hi !== -1) {
    const anchors = lines[hi].cells.map((c) => ({ x: (c.x + c.xEnd) / 2, n: gnorm(c.str) }));
    const gIdx = anchors.findIndex((a) => a.n.includes("grade"));
    const wIdx = anchors.findIndex((a) => a.n.includes("totalweight"));
    for (let i = hi + 1; i < lines.length && gIdx !== -1 && wIdx !== -1; i++) {
      const assign = {};
      for (const c of lines[i].cells) {
        let bi = -1, bd = 1e9;
        anchors.forEach((a, ai) => {
          const d = Math.abs((c.x + c.xEnd) / 2 - a.x);
          if (d < bd) { bd = d; bi = ai; }
        });
        assign[bi] = assign[bi] ? assign[bi] + " " + c.str : c.str;
      }
      const label = gnorm(assign[gIdx] || "");
      if (label === "total") break; // end of GRADE table; SIZE/PACKOUT tables are ignored
      const f = GRADE_MAP[label];
      if (f && assign[wIdx] != null) rec[f] = num(String(assign[wIdx]).replace(/,/g, ""));
    }
  }
  return { sheet: "Classification report", rows: [rec] };
}

async function parsePdfFile(file) {
  const pdfjs = await loadPdfJs();
  const buf = await file.arrayBuffer();
  const doc = await pdfjs.getDocument({ data: buf }).promise;
  const lines = [];
  let anyText = false;
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    const items = tc.items.map((it) => ({ x: it.transform[4], y: it.transform[5], w: it.width, str: it.str }));
    if (items.some((it) => it.str && it.str.trim())) anyText = true;
    pdfLines(items).forEach((l) => lines.push(l));
  }
  if (!anyText) throw new Error("pdf-noText");
  return pdfRecords(lines) || pdfClassification(lines);
}

function ImportCard({ recs, onImportMany, defaultBw, ready }) {
  const [parsed, setParsed] = useState(null);
  const [bws, setBws] = useState([]); // per-line Av. berry weight entries
  const [fileName, setFileName] = useState("");
  const [err, setErr] = useState("");
  const [done, setDone] = useState(null);
  const [busy, setBusy] = useState(false);

  const existing = useMemo(() => new Set(recs.map((r) => (r.date || "") + "|" + (r.batch || ""))), [recs]);
  const isDup = (r) => existing.has(r.date + "|" + (r.batch || ""));

  const onFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = ""; // allow re-selecting the same file
    if (!file) return;
    setErr(""); setDone(null); setParsed(null); setFileName(file.name); setBusy(true);
    const isPdf = /\.pdf$/i.test(file.name) || file.type === "application/pdf";
    if (isPdf && file.size > 30 * 1024 * 1024) {
      setErr("This PDF is " + (file.size / 1048576).toFixed(0) + " MB \u2014 too large to read comfortably on-device. Export fewer pages and upload in parts.");
      setBusy(false);
      return;
    }
    try {
      const res = isPdf
        ? await parsePdfFile(file)
        : parseBatchFile(new Uint8Array(await file.arrayBuffer()));
      if (!res) {
        setErr(isPdf
          ? "No batch records could be found in this PDF."
          : "Couldn't find batch data in this file. It needs a header row containing \"Date Processed\" and \"Batch Number\" (like the Processed Data Entry sheet).");
      } else {
        setParsed(res);
        // file value wins; otherwise the default entered before upload
        setBws(res.rows.map((r) => (r.bw != null ? String(r.bw) : String(defaultBw))));
      }
    } catch (ex) {
      const msg = (ex && ex.message) || "";
      if (msg === "pdf-lib-net") setErr("Couldn't download the PDF reader component \u2014 the first PDF import on a device needs internet. Check the connection and retry (afterwards it's cached).");
      else if (msg === "pdf-noText") setErr("This PDF has no readable text layer \u2014 it looks like a scan or photo. On-device extraction needs a text PDF: export or print-to-PDF from Excel instead of scanning paper. (Or upload the .xlsx/.csv directly.)");
      else setErr("Couldn't read this file (" + ((ex && ex.message) || "unknown error").slice(0, 120) + "). Supported: .pdf, .xlsx, .xlsm, .csv \u2014 screenshot this and send it to Claude if it persists.");
    }
    setBusy(false);
  };

  const dupes = parsed ? parsed.rows.filter(isDup) : [];
  const freshIdx = parsed ? parsed.rows.map((r, i) => (isDup(r) ? -1 : i)).filter((i) => i >= 0) : [];
  const missing = freshIdx.filter((i) => !num(bws[i])).length;

  const doImport = async () => {
    setBusy(true);
    // each line carries its own Av. berry weight
    const withBw = freshIdx.map((i) => ({ ...parsed.rows[i], bw: num(bws[i]) }));
    const ok = await onImportMany(withBw);
    setBusy(false);
    setDone(ok
      ? "Submitted " + freshIdx.length + " batch" + (freshIdx.length === 1 ? "" : "es") + (dupes.length ? " (" + dupes.length + " duplicates skipped)" : "") + "."
      : "Submitted for this session, but shared storage is unavailable — it may not persist.");
    setParsed(null);
  };

  return (
    <Card title="Import from file" right={<span className="muted small">.pdf · .xlsx · .xlsm · .csv</span>}>
      <p className="muted" style={{ marginBottom: 10 }}>
        Upload a PDF report or a workbook and the batch records will be extracted automatically, entirely on this device — no key or AI service needed. Supported PDFs: TrueSort / Ellips Classification Reports (one batch per file) and Processed Data Entry sheet printouts — text-based only, scanned paper can’t be read. Every line is pre-filled with the berry weight entered above (file values win) and stays editable before you submit all lines at once. Duplicates already in the register are skipped. One file at a time.
      </p>
      {!ready && <p className="savedmsg warn" style={{ marginBottom: 10 }}>Enter the Av. berry weight above before choosing a file.</p>}
      <input type="file" accept=".pdf,.xlsx,.xlsm,.xls,.csv" onChange={onFile} className="filein" disabled={!ready || busy || !!parsed} />
      {busy && <p className="muted" style={{ marginTop: 10 }}>Reading… PDF extraction can take 10–30 seconds.</p>}
      {err && <p className="savedmsg warn" style={{ marginTop: 10 }}>{err}</p>}
      {done && <p className="savedmsg ok" style={{ marginTop: 10 }}>{done}</p>}
      {parsed && (
        <div style={{ marginTop: 12 }}>
          <p style={{ marginBottom: 8 }}>
            <b>{fileName}</b> · sheet "{parsed.sheet}" · found <b>{parsed.rows.length}</b> batch record{parsed.rows.length === 1 ? "" : "s"}
            {dupes.length > 0 && <span className="neg"> · {dupes.length} already in the register (will be skipped)</span>}
          </p>

          <div className="tbl-wrap" style={{ maxHeight: 220, overflowY: "auto", border: "1px solid var(--line)", borderRadius: 8 }}>
            <table>
              <thead><tr><th>Date</th><th>Batch</th><th>Block</th><th>Variety</th><th className="r">Processed kg</th><th className="r">Berry wt g</th><th></th></tr></thead>
              <tbody>
                {parsed.rows.slice(0, 200).map((r, i) => {
                  const dup = isDup(r);
                  return (
                    <tr key={i} className={dup ? "hidef" : ""}>
                      <td>{dfmt(r.date)}</td><td className="mono">{r.batch || "—"}</td>
                      <td>{r.block || "—"}</td><td>{r.variety || "—"}</td>
                      <td className="r">{kg(r.processed, 1)}</td>
                      <td className="r">
                        {dup ? (r.bw ?? "—") : (
                          <input className="inp bw-in" type="number" step="any" min="0"
                            value={bws[i] ?? ""} placeholder="req."
                            onChange={(e) => { const a = [...bws]; a[i] = e.target.value; setBws(a); }} />
                        )}
                      </td>
                      <td className="muted small">{dup ? "duplicate" : ""}</td>
                    </tr>
                  );
                })}
                {parsed.rows.length > 200 && <tr><td colSpan={7} className="muted center">…and {parsed.rows.length - 200} more</td></tr>}
              </tbody>
            </table>
          </div>
          {missing > 0 && <p className="savedmsg warn" style={{ marginTop: 8 }}>{missing} line{missing === 1 ? " still needs" : "s still need"} a berry weight before submitting.</p>}
          <div className="oed-actions">
            <button className="btn btn-primary" style={{ width: "auto" }} disabled={!freshIdx.length || missing > 0 || busy} onClick={doImport}>
              Submit {freshIdx.length} batch{freshIdx.length === 1 ? "" : "es"}
            </button>
            <button className="btn btn-mini" onClick={() => setParsed(null)}>Cancel</button>
          </div>
        </div>
      )}
    </Card>
  );
}

function NewBatch({ recs, onSaveMany, onDeleteApp, me }) {
  const [defBw, setDefBw] = useState("");
  const [confirmDel, setConfirmDel] = useState(false);
  const [delMsg, setDelMsg] = useState(null);

  const appBatches = recs.filter((r) => r.source === "app");
  const deletable = me.role === "admin" ? appBatches : []; // deleting uploads is admin-only

  const doDelete = async () => {
    const ok = await onDeleteApp();
    setConfirmDel(false);
    setDelMsg(ok
      ? "Deleted " + deletable.length + " batch" + (deletable.length === 1 ? "" : "es") + ". You can now re-upload the corrected file."
      : "Deleted for this session, but shared storage is unavailable — it may not persist.");
  };

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Batch upload</h2>
          <p className="muted">Enter the Av. berry weight, upload one file at a time, adjust any line that differs, then submit all lines together.</p>
        </div>
      </div>

      <Card title="Av. berry weight (required before upload)">
        <div className="fgrid" style={{ maxWidth: 340 }}>
          <label className="fld"><span>Av. berry weight (g)</span>
            <input className="inp" type="number" step="any" min="0" value={defBw} onChange={(e) => setDefBw(e.target.value)} />
          </label>
        </div>
        <p className="muted small" style={{ marginTop: 8 }}>
          {num(defBw) ? "Pre-fills every extracted line — each line stays editable before submitting." : "Required before a file can be uploaded."}
        </p>
      </Card>

      <ImportCard recs={recs} onImportMany={onSaveMany} defaultBw={num(defBw)} ready={!!num(defBw)} />

      {me.role === "admin" && (
      <Card title="Delete & reload" right={<span className="muted small">{appBatches.length} app-entered batch{appBatches.length === 1 ? "" : "es"} in the register</span>}>
        <p className="muted" style={{ marginBottom: 10 }}>
          If an upload came in wrong, delete the app-entered batches and upload the file again.
          This removes them from the shared register for everyone — re-import afterwards. Deleting is admin-only.
        </p>
        {!confirmDel ? (
          <button className="btn btn-mini btn-danger" disabled={!deletable.length} onClick={() => { setConfirmDel(true); setDelMsg(null); }}>
            Delete {deletable.length} batch{deletable.length === 1 ? "" : "es"}…
          </button>
        ) : (
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <span style={{ fontWeight: 600, color: "var(--red)" }}>This can't be undone. Sure?</span>
            <button className="btn btn-mini btn-danger" onClick={doDelete}>Yes, delete</button>
            <button className="btn btn-mini" onClick={() => setConfirmDel(false)}>Cancel</button>
          </div>
        )}
        {delMsg && <p className="savedmsg ok" style={{ marginTop: 10 }}>{delMsg}</p>}
      </Card>
      )}
    </>
  );
}

/* ================= QUALITY ================= */
function Quality({ recs }) {
  const dates = [...new Set(recs.filter((r) => num(r.processed)).map((r) => r.date))].sort().reverse(); // newest first
  const [date, setDate] = useState("");
  const scope = recs.filter((r) => num(r.processed) && (!date || r.date === date));
  const totP = scope.reduce((s, r) => s + n0(r.processed), 0);
  const pctOf = (kgv, base) => (base ? (kgv / base) * 100 : 0);

  // headline strip (matches the workbook dashboard)
  const headKeys = [["tears", "Tear %"], ["redBums", "Red bums %"], ["greenBums", "Green bums %"],
    ["shrivell", "Shrivell %"], ["subBloom", "Sub bloom %"], ["softPoint", "Soft point %"]];
  const totalDef = pctOf(scope.reduce((s, r) => s + derive(r).defKg, 0), totP);

  // variety + block rows
  const byVB = {};
  for (const r of scope) {
    const k = (r.variety || "(blank)") + "|" + (r.block || "(blank)");
    (byVB[k] ||= []).push(r);
  }
  const rows = Object.entries(byVB).map(([k, rs]) => {
    const [variety, block] = k.split("|");
    const p = rs.reduce((t, r) => t + n0(r.processed), 0);
    const bwW = rs.filter((r) => num(r.bw) && num(r.processed));
    const bwT = bwW.reduce((t, r) => t + r.processed, 0);
    return {
      variety, block, p,
      bw: bwT ? bwW.reduce((t, r) => t + r.bw * r.processed, 0) / bwT : null,
      def: DEFECTS.map(([key]) => pctOf(rs.reduce((t, r) => t + n0(r[key]), 0), p)),
    };
  }).sort((a, b) => a.variety.localeCompare(b.variety) || a.block.localeCompare(b.block));

  const totalRow = {
    bw: (() => {
      const w = scope.filter((r) => num(r.bw) && num(r.processed));
      const t = w.reduce((x, r) => x + r.processed, 0);
      return t ? w.reduce((x, r) => x + r.bw * r.processed, 0) / t : null;
    })(),
    def: DEFECTS.map(([key]) => pctOf(scope.reduce((t, r) => t + n0(r[key]), 0), totP)),
  };

  // charts: total defect % by variety, and per-defect averages
  const byVar = {};
  for (const r of scope) {
    const v = r.variety || "(blank)";
    (byVar[v] ||= { p: 0, def: 0 });
    byVar[v].p += n0(r.processed);
    byVar[v].def += derive(r).defKg;
  }
  const varChart = Object.entries(byVar)
    .map(([name, x]) => ({ name, pct: +pctOf(x.def, x.p).toFixed(1) }))
    .sort((a, b) => a.name.localeCompare(b.name));
  const typeChart = DEFECTS.map(([key, label], i) => ({ name: label, pct: +totalRow.def[i].toFixed(2) }));

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Quality Defects</h2>
          <p className="muted">{date ? dlong(date) + " \u00b7 Wk " + isoWeek(date) : "Whole season"}</p>
        </div>
        <div className="controls">
          <label className="lbl">Date</label>
          <Select value={date} onChange={setDate} options={dates} all="Whole season" />
          <PdfBtn name="Quality_Defects" />
        </div>
      </div>

      <div className="kpis">
        <Kpi label="Total defect %" value={totalDef.toFixed(1) + "%"} tone={totalDef > 10 ? "bad" : undefined} sub={"\u00A0"} />
        {headKeys.map(([key, label]) => (
          <Kpi key={key} label={label}
            value={pctOf(scope.reduce((t, r) => t + n0(r[key]), 0), totP).toFixed(1) + "%"} sub={"\u00A0"} />
        ))}
      </div>

      <Card title="Defects by variety and block" pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>Variety</th><th>Block</th><th className="r">Av. berry wt</th>
                {DEFECTS.map(([, l]) => <th key={l} className="r">{l} %</th>)}
              </tr>
            </thead>
            <tbody>
              {rows.map((x) => (
                <tr key={x.variety + x.block}>
                  <td className="strong">{x.variety}</td><td>{x.block}</td>
                  <td className="r">{x.bw != null ? x.bw.toFixed(1) : "\u2014"}</td>
                  {x.def.map((v, i) => <td key={i} className="r">{v.toFixed(1)}%</td>)}
                </tr>
              ))}
              {!rows.length && <tr><td colSpan={14} className="muted center">No batches in scope.</td></tr>}
              <tr className="totalrow">
                <td className="strong" colSpan={2}>TOTAL (all blocks)</td>
                <td className="r strong">{totalRow.bw != null ? totalRow.bw.toFixed(1) : "\u2014"}</td>
                {totalRow.def.map((v, i) => <td key={i} className="r strong">{v.toFixed(1)}%</td>)}
              </tr>
            </tbody>
          </table>
        </div>
      </Card>

      <div className="grid2">
        <Card title="Total defect % by variety">
          <SBarChart data={varChart} xKey="name" height={240} unit="%"
            series={[{ key: "pct", label: "Total defect", color: "#4472C4" }]} />
        </Card>

        <Card title={"Defect types \u2014 " + (date ? "day" : "season") + " avg %"}>
          <SBarChart data={typeChart} xKey="name" height={240} unit="%" angled
            series={[{ key: "pct", label: "Share", color: "#4472C4" }]} />
        </Card>
      </div>
    </>
  );
}

/* ================= SEASON ================= */
function exportSeasonReport(recs, me) {
  const data = recs.filter((r) => num(r.processed));
  if (!data.length) return false;
  const dates = data.map((r) => r.date).filter(Boolean).sort();
  const withD = data.map((r) => ({ r, d: derive(r) }));
  const sum = (fn) => withD.reduce((t, x) => t + (fn(x) || 0), 0);
  const P = sum((x) => n0(x.r.processed));
  const pct = (v) => (P ? +((v / P) * 100).toFixed(2) : 0);

  const wb = XLSX.utils.book_new();
  const aoa = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);

  // --- Summary ---
  aoa("Summary", [
    ["MERRYLIGHT PACKHOUSE \u2014 SEASON REPORT"],
    ["Generated", new Date().toLocaleString("en-GB")],
    ["Generated by", me ? me.name + " (" + (ROLE_LABEL[me.role] || me.role) + ")" : ""],
    ["Date range", dfmt(dates[0]) + " \u2013 " + dfmt(dates[dates.length - 1])],
    [],
    ["Batches", data.length],
    ["Blocks", new Set(data.map((r) => r.block).filter(Boolean)).size],
    ["Varieties", new Set(data.map((r) => r.variety).filter(Boolean)).size],
    ["Processed kg", +P.toFixed(1)],
    ["Reaped kg", +sum((x) => n0(x.r.reaped)).toFixed(1)],
    [],
    ["SIZE MIX", "% of processed"],
    ["12mm+", pct(sum((x) => x.d.k12))], ["14mm+", pct(sum((x) => x.d.k14))],
    ["18mm+", pct(sum((x) => x.d.k18))], ["22mm+", pct(sum((x) => x.d.k22))],
    ["<12mm", pct(sum((x) => n0(x.r.under12)))],
    [],
    ["FREIGHT", "% of processed"],
    ["Air", pct(sum((x) => x.d.air))], ["Sea", pct(sum((x) => x.d.sea))],
    [],
    ["QUALITY", "% of processed"],
    ...Object.entries(GROUPS).map(([g, keys]) => [g, pct(sum((x) => keys.reduce((t, k) => t + n0(x.r[k]), 0)))]),
    ["Total defects", pct(sum((x) => x.d.defKg))],
  ]);

  // --- By variety ---
  const groupBy = (keyFn) => {
    const m = {};
    for (const x of withD) {
      const k = keyFn(x.r) || "(blank)";
      (m[k] ||= []).push(x);
    }
    return m;
  };
  const rollRows = (m, label) => {
    const out = [[label, "Batches", "Processed kg", "12mm+ %", "14mm+ %", "18mm+ %", "22mm+ %", "Air %", "Sea %", "Defect %"]];
    for (const [k, xs] of Object.entries(m).sort((a, b) => b[1].reduce((t, x) => t + n0(x.r.processed), 0) - a[1].reduce((t, x) => t + n0(x.r.processed), 0))) {
      const p = xs.reduce((t, x) => t + n0(x.r.processed), 0);
      const pc2 = (v) => (p ? +((v / p) * 100).toFixed(1) : 0);
      out.push([k, xs.length, +p.toFixed(1),
        pc2(xs.reduce((t, x) => t + x.d.k12, 0)), pc2(xs.reduce((t, x) => t + x.d.k14, 0)),
        pc2(xs.reduce((t, x) => t + x.d.k18, 0)), pc2(xs.reduce((t, x) => t + x.d.k22, 0)),
        pc2(xs.reduce((t, x) => t + x.d.air, 0)), pc2(xs.reduce((t, x) => t + x.d.sea, 0)),
        pc2(xs.reduce((t, x) => t + x.d.defKg, 0))]);
    }
    return out;
  };
  aoa("By variety", rollRows(groupBy((r) => r.variety), "Variety"));
  aoa("By block", rollRows(groupBy((r) => r.block), "Block"));
  aoa("By week", rollRows(groupBy((r) => (r.date ? "Wk " + String(isoWeek(r.date)).padStart(2, "0") : null)), "Week"));

  // --- Full batch register ---
  const reg = [["Date", "Week", "Batch", "Block", "Variety", "Section", "Reaped kg", "Processed kg", "Berry wt g",
    "12+Air", "14+Air", "18+Air", "22+Air", "12+Sea", "14+Sea", "18+Sea", "22+Sea",
    ...DEFECTS.map(([, l]) => l), "Defect %", "Entered by", "Source"]];
  for (const { r, d } of withD.sort((a, b) => (a.r.date || "").localeCompare(b.r.date || ""))) {
    reg.push([r.date, r.date ? isoWeek(r.date) : null, r.batch, r.block, r.variety, r.section,
      r.reaped, r.processed, r.bw,
      r.a12, r.a14, r.a18, r.a22, r.s12, r.s14, r.s18, r.s22,
      ...DEFECTS.map(([k]) => r[k]),
      d.defPct != null ? +(d.defPct * 100).toFixed(2) : null,
      r.enteredBy || null, r.source || "workbook"]);
  }
  aoa("Batch register", reg);

  XLSX.writeFile(wb, "Merrylight_Season_Report_" + dates[dates.length - 1] + ".xlsx");
  return true;
}

function Season({ recs, me }) {
  const weeks = [...new Set(recs.filter((r) => num(r.processed) && r.date).map((r) => isoWeek(r.date)))].sort((a, b) => b - a); // newest first
  const [week, setWeek] = useState("");
  const data = recs.filter((r) => num(r.processed) && r.variety && !["ALL", "REJECTS"].includes(r.variety)
    && (!week || (r.date && isoWeek(r.date) === +week)));
  const byVar = {};
  for (const r of data) {
    const d = derive(r);
    const v = (byVar[r.variety] ||= { variety: r.variety, kg: 0, reaped: 0, k12: 0, k14: 0, k18: 0, k22: 0, u12: 0, def: 0, air: 0, sea: 0, n: 0, blocks: new Set() });
    v.kg += r.processed; v.reaped += n0(r.reaped); v.n++;
    v.k12 += d.k12; v.k14 += d.k14; v.k18 += d.k18; v.k22 += d.k22;
    v.u12 += n0(r.under12); v.def += d.defKg; v.air += d.air; v.sea += d.sea;
    v.blocks.add(r.block);
  }
  const rows = Object.values(byVar).sort((a, b) => b.kg - a.kg);
  const totKg = rows.reduce((s, v) => s + v.kg, 0);
  const chart = rows.map((v) => ({
    name: v.variety,
    "14mm+": +((v.k14 / (v.kg || 1)) * 100).toFixed(1),
    "18mm+": +((v.k18 / (v.kg || 1)) * 100).toFixed(1),
    "22mm+": +((v.k22 / (v.kg || 1)) * 100).toFixed(1),
    kg: +v.kg.toFixed(0),
  }));

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Season summary</h2>
          <p className="muted">{week ? "Week " + week : "Whole season"} · named varieties only — "ALL" and "REJECTS" lines excluded · {kg(totKg)} kg processed</p>
        </div>
        <div className="controls">
          <label className="lbl">Week</label>
          <Select value={week} onChange={setWeek} options={weeks.map(String)} all="Whole season" />
          <button className="btn btn-primary" style={{ width: "auto", padding: "7px 13px" }}
            onClick={() => exportSeasonReport(recs, me)}>
            Export season report
          </button>
          <PdfBtn name="Season" />
        </div>
      </div>

      <Card title="Volume and premium sizes by variety" right={<span className="muted small">bars: size share · label: total kg</span>}>
        <SBarChart data={chart} xKey="name" height={280} unit="%" stacked
          series={[
            { key: "14mm+", color: SIZE_COLORS.k14 },
            { key: "18mm+", color: SIZE_COLORS.k18 },
            { key: "22mm+", color: SIZE_COLORS.k22 },
          ]} />
      </Card>

      <Card title="Variety league table" pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>Variety</th><th className="r">Batches</th><th className="r">Blocks</th>
                <th className="r">Processed kg</th><th className="r">Share</th>
                <th className="r">12mm+</th><th className="r">14mm+</th><th className="r">18mm+</th><th className="r">22mm+</th>
                <th className="r">Air / Sea</th><th className="r">Defects</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((v) => {
                return (
                  <tr key={v.variety} className={v.def / v.kg > 0.1 ? "hidef" : ""}>
                    <td className="strong">{v.variety}</td>
                    <td className="r">{v.n}</td>
                    <td className="r">{v.blocks.size}</td>
                    <td className="r strong">{kg(v.kg)}</td>
                    <td className="r">{pc(v.kg / totKg)}</td>
                    
                    <td className="r">{pc(v.k12 / v.kg)}</td><td className="r">{pc(v.k14 / v.kg)}</td><td className="r">{pc(v.k18 / v.kg)}</td><td className="r">{pc(v.k22 / v.kg)}</td>
                    <td className="r">{v.air + v.sea ? pc(v.air / (v.air + v.sea), 0) + " / " + pc(v.sea / (v.air + v.sea), 0) : "—"}</td>
                    <td className={"r" + (v.def / v.kg > 0.1 ? " neg strong" : "")}>{pc(v.def / v.kg)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </>
  );
}

/* ================= LOGIN ================= */
function Login({ onLogin, onReset, firstRun }) {
  const [be, setBe] = useState(null); // backend status from the hosted storage layer
  useEffect(() => {
    if (typeof window === "undefined" || !window.__mlBackend) return;
    const t = setInterval(() => {
      setBe({ ...window.__mlBackend });
      if (window.__mlBackend.status !== "untested") clearInterval(t);
    }, 600);
    return () => clearInterval(t);
  }, []);
  const [help, setHelp] = useState(false);
  const [confirmReset, setConfirmReset] = useState(false);
  const [resetMsg, setResetMsg] = useState(null);
  const [u, setU] = useState("");
  const [p, setP] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const submit = async (e) => {
    if (e) e.preventDefault();
    if (busy) return;
    setBusy(true); setErr("");
    try {
      const bad = await onLogin(u, p);
      if (bad) { setErr(bad); setBusy(false); }
    } catch {
      setErr("Sign-in failed unexpectedly. Try again, or use 'Trouble signing in?' below.");
      setBusy(false);
    }
  };
  return (
    <div className="login-wrap">
      <div className="login-card">
        <div className="brand login-brand">
          <img className="brand-logo" src={LOGO} alt="Merrylight Enterprises" />
          <div>
            <div className="brand-name" style={{ color: "var(--ink)" }}>MERRYLIGHT</div>
            <div className="brand-sub">Packhouse ERP · Season 2026</div>
          </div>
        </div>
        <label className="fld"><span>Username</span>
          <input className="inp" autoFocus value={u} onChange={(e) => { setU(e.target.value); setErr(""); }}
            onKeyDown={(e) => e.key === "Enter" && submit()}
            autoComplete="off" name="ml-u" autoCapitalize="none" autoCorrect="off" spellCheck={false} />
        </label>
        <label className="fld"><span>PIN</span>
          <input className="inp" type="password" value={p} onChange={(e) => { setP(e.target.value); setErr(""); }}
            onKeyDown={(e) => e.key === "Enter" && submit()} autoComplete="new-password" name="ml-p" />
        </label>
        {err && <p className="login-err">{err}</p>}
        <button className="btn btn-primary" disabled={busy || !u.trim() || !p} onClick={submit}>
          {busy ? "Signing in…" : "Sign in"}
        </button>
        {firstRun && (
          <p className="login-hint">First run: sign in as <b>admin</b> with PIN <b>1234</b>, then change it on the Users page.</p>
        )}
        {be && (
          <p className="login-note" style={be.status === "error" ? { color: "var(--red)", fontWeight: 600 } : be.status === "off" ? { color: "var(--amber)" } : { color: "var(--leaf)" }}>
            Team backend: {be.status === "ok" ? "connected ✓"
              : be.status === "off" ? "not configured — device-only mode (accounts and data stay on this device)"
              : be.status === "untested" ? "checking…"
              : "ERROR — " + (be.detail || "unreachable") + " — screenshot this and send it to Claude."}
          </p>
        )}
        <button className="linklike" onClick={() => { setHelp(!help); setConfirmReset(false); setResetMsg(null); }}>Trouble signing in?</button>
        {help && (
          <div className="login-help">
            <p>Default first-run sign-in is <b>admin</b> / <b>1234</b>. If that's been changed and no one knows a working PIN, you can reset the accounts. This restores sign-in to admin / 1234 only — batch records, orders, and stock lines are untouched.</p>
            {!confirmReset ? (
              <button className="btn btn-mini btn-danger" onClick={() => setConfirmReset(true)}>Reset accounts…</button>
            ) : (
              <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <span style={{ fontWeight: 600, color: "var(--red)" }}>This removes all user accounts. Sure?</span>
                <button className="btn btn-mini btn-danger" onClick={async () => {
                  await onReset();
                  setConfirmReset(false);
                  setResetMsg("Accounts reset — sign in with admin / 1234, then set a new PIN and re-add your team on the Users page.");
                }}>Yes, reset</button>
                <button className="btn btn-mini" onClick={() => setConfirmReset(false)}>Cancel</button>
              </div>
            )}
            {resetMsg && <p className="savedmsg ok">{resetMsg}</p>}
          </div>
        )}
      </div>
    </div>
  );
}

/* ================= USERS ================= */
const PAGE_OPTIONS = [
  ["dash", "Dashboard"], ["batches", "Batches"], ["new", "Batch upload"],
  ["quality", "Quality Defects"], ["season", "Season"], ["packshed", "Packshed"], ["chat", "Chat"],
  ["compliance", "Compliance"],
  ["hr", "Human Resources"],
]; // "users" is always admin-only and not assignable

function Users({ me, users, onChange }) {
  const [form, setForm] = useState({ name: "", username: "", role: "entry", pin: "", pages: PAGE_OPTIONS.map(([k]) => k) });
  const qcDefault = ["compliance"];
  const hrDefault = ["hr"];
  const [msg, setMsg] = useState(null);
  const [pinCur, setPinCur] = useState("");
  const [pinNew, setPinNew] = useState("");
  const [pinMsg, setPinMsg] = useState(null);
  const setF = (k) => (e) => { setForm({ ...form, [k]: e.target.value }); setMsg(null); };

  const admins = users.filter((x) => x.role === "admin").length;

  const [editFor, setEditFor] = useState(null);
  const [draft, setDraft] = useState(null); // { name, role, pages }
  const openEdit = (u) => {
    if (editFor === u.username) return setEditFor(null);
    setEditFor(u.username);
    setDraft({ name: u.name, role: u.role, pages: u.pages || PAGE_OPTIONS.map(([k]) => k) });
    setMsg(null);
  };
  const saveEdit = async (u) => {
    const name = String(draft.name || "").trim();
    if (!name) return setMsg(["warn", "Name is required."]);
    if (u.username === me.username && draft.role !== u.role)
      return setMsg(["warn", "You can't change your own role — ask another admin."]);
    if (u.role === "admin" && draft.role !== "admin" && admins <= 1)
      return setMsg(["warn", "At least one admin must remain."]);
    if (draft.role !== "admin" && !draft.pages.length) return setMsg(["warn", "Select at least one page."]);
    const pages = draft.role === "admin" || draft.pages.length === PAGE_OPTIONS.length ? null : draft.pages;
    const ok = await onChange(users.map((x) => (x.username === u.username ? { ...x, name, role: draft.role, pages } : x)));
    setEditFor(null);
    setMsg(ok ? ["ok", "Updated " + u.username + "."] : ["warn", "Saved for this session, but shared storage is unavailable."]);
  };

  const add = async () => {
    const name = form.name.trim(), username = form.username.trim().toLowerCase();
    if (!name || !username) return setMsg(["warn", "Name and username are required."]);
    if (!/^[a-z0-9._-]{2,20}$/.test(username)) return setMsg(["warn", "Username: 2–20 characters, letters, numbers, . _ - only."]);
    if (users.some((x) => x.username === username)) return setMsg(["warn", "That username is taken."]);
    if (form.pin.length < 4) return setMsg(["warn", "PIN must be at least 4 characters."]);
    if (form.role !== "admin" && !form.pages.length) return setMsg(["warn", "Select at least one page they can see."]);
    const pages = form.role === "admin" || form.pages.length === PAGE_OPTIONS.length ? null : form.pages;
    const rec = { name, username, role: form.role, pin: await sha(form.pin), created: Date.now(), pages };
    const ok = await onChange([...users, rec]);
    setMsg(ok ? ["ok", "Added " + name + " (" + username + ")."] : ["warn", "Added for this session, but shared storage is unavailable."]);
    setForm({ name: "", username: "", role: "entry", pin: "", pages: PAGE_OPTIONS.map(([k]) => k) });
  };

  const remove = async (u) => {
    if (u.username === me.username) return setMsg(["warn", "You can't remove your own account."]);
    if (u.role === "admin" && admins <= 1) return setMsg(["warn", "At least one admin must remain."]);
    await onChange(users.filter((x) => x.username !== u.username));
    setMsg(["ok", "Removed " + u.username + "."]);
  };

  const resetPin = async (u) => {
    const pin = String(Math.floor(1000 + Math.random() * 9000));
    const hashed = await sha(pin);
    await onChange(users.map((x) => (x.username === u.username ? { ...x, pin: hashed } : x)));
    setMsg(["ok", "New PIN for " + u.username + ": " + pin + " — note it down now, it won't be shown again."]);
  };

  const changeMyPin = async () => {
    if ((await sha(pinCur)) !== me.pin) return setPinMsg(["warn", "Current PIN is incorrect."]);
    if (pinNew.length < 4) return setPinMsg(["warn", "New PIN must be at least 4 characters."]);
    const hashed = await sha(pinNew);
    await onChange(users.map((x) => (x.username === me.username ? { ...x, pin: hashed } : x)));
    setPinCur(""); setPinNew("");
    setPinMsg(["ok", "PIN updated."]);
  };

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Users</h2>
          <p className="muted">{users.length} account{users.length === 1 ? "" : "s"} · shared across everyone using this app</p>
        </div>
      </div>

      <Card title="Team" pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr><th>Name</th><th>Username</th><th>Role</th><th>Access</th><th className="r">Actions</th></tr>
            </thead>
            <tbody>
              {users.map((u) => (
                <React.Fragment key={u.username}>
                <tr>
                  <td className="strong">{u.name}{u.username === me.username && <span className="badge">you</span>}</td>
                  <td className="mono">{u.username}</td>
                  <td>{ROLE_LABEL[u.role] || u.role}</td>
                  <td className="muted small">
                    {u.role === "admin" ? "All pages (admin)"
                      : !u.pages ? "All pages"
                      : u.pages.map((k) => (PAGE_OPTIONS.find(([p]) => p === k) || [k, k])[1]).join(", ")}
                  </td>
                  <td className="r">
                    <button className="btn btn-mini" onClick={() => openEdit(u)}>
                      {editFor === u.username ? "Close" : "Edit"}
                    </button>{" "}
                    <button className="btn btn-mini" onClick={() => resetPin(u)}>Reset PIN</button>{" "}
                    <button className="btn btn-mini btn-danger" disabled={u.username === me.username || (u.role === "admin" && admins <= 1)}
                      onClick={() => remove(u)}>Remove</button>
                  </td>
                </tr>
                {editFor === u.username && draft && (
                  <tr className="detail"><td colSpan={5}>
                    <div style={{ padding: "10px 4px", position: "sticky", left: 12, maxWidth: "min(660px, calc(100vw - 120px))" }}>
                      <div className="fgrid">
                        <label className="fld"><span>Full name</span>
                          <input className="inp" value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} />
                        </label>
                        <label className="fld"><span>Username (fixed)</span>
                          <input className="inp" value={u.username} disabled />
                        </label>
                        <label className="fld"><span>Role</span>
                          <select className="inp" value={draft.role} disabled={u.username === me.username}
                            onChange={(e) => setDraft({ ...draft, role: e.target.value })}>
                            <option value="admin">Admin — everything, incl. users & deletions</option>
                            <option value="entry">Data entry — add batches</option>
                            <option value="viewer">View only</option>
                          
                            <option value="qc">QC — compliance forms only</option>
                            <option value="hrOfficer">HR Officer — fill &amp; submit HR forms</option>
                            <option value="hrManager">HR Manager — full HR access including contract edits</option>
                          </select>
                        </label>
                      </div>
                      {draft.role !== "admin" && (
                        <div style={{ marginTop: 10 }}>
                          <p className="muted small" style={{ marginBottom: 8 }}>Pages {draft.name || u.username} can see:</p>
                          <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
                            {PAGE_OPTIONS.map(([k, label]) => (
                              <label key={k} style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 12.5, cursor: "pointer" }}>
                                <input type="checkbox" checked={draft.pages.includes(k)}
                                  onChange={() => setDraft({ ...draft, pages: draft.pages.includes(k) ? draft.pages.filter((x) => x !== k) : [...draft.pages, k] })} />
                                {label}
                              </label>
                            ))}
                          </div>
                        </div>
                      )}
                      <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                        <button className="btn btn-mini btn-primary" onClick={() => saveEdit(u)}>Save changes</button>
                        <button className="btn btn-mini" onClick={() => setEditFor(null)}>Cancel</button>
                      </div>
                    </div>
                  </td></tr>
                )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
        {msg && <p className={"savedmsg " + msg[0]} style={{ padding: "0 16px 14px" }}>{msg[1]}</p>}
      </Card>

      <div className="grid2">
        <Card title="Add user">
          <div className="fgrid">
            <label className="fld"><span>Full name</span><input className="inp" value={form.name} onChange={setF("name")} /></label>
            <label className="fld"><span>Username</span><input className="inp" value={form.username} onChange={setF("username")} /></label>
            <label className="fld"><span>Role</span>
              <select className="inp" value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value, pages: e.target.value === "qc" ? ["compliance"] : (e.target.value === "hrOfficer" || e.target.value === "hrManager") ? ["hr"] : PAGE_OPTIONS.map(([k]) => k) })}>
                <option value="admin">Admin — everything, incl. users & deletions</option>
                <option value="entry">Data entry — add batches</option>
                <option value="viewer">View only</option>
              
                <option value="qc">QC — compliance forms only</option>
                <option value="hrOfficer">HR Officer — fill &amp; submit HR forms</option>
                <option value="hrManager">HR Manager — full HR access including contract edits</option>
              </select>
            </label>
            <label className="fld"><span>PIN (min 4 chars)</span><input className="inp" type="password" value={form.pin} onChange={setF("pin")} autoComplete="new-password" /></label>
          </div>
          {form.role !== "admin" && (
            <div style={{ marginTop: 12 }}>
              <p className="muted small" style={{ marginBottom: 8 }}>Page access — untick what they shouldn't see:</p>
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
                {PAGE_OPTIONS.map(([k, label]) => (
                  <label key={k} style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 12.5, cursor: "pointer" }}>
                    <input type="checkbox" checked={form.pages.includes(k)}
                      onChange={() => setForm({ ...form, pages: form.pages.includes(k) ? form.pages.filter((x) => x !== k) : [...form.pages, k] })} />
                    {label}
                  </label>
                ))}
              </div>
            </div>
          )}
          <button className="btn btn-primary" style={{ marginTop: 12, width: "auto" }} onClick={add}>Add user</button>
        </Card>

        <Card title="Change my PIN">
          <div className="fgrid">
            <label className="fld"><span>Current PIN</span><input className="inp" type="password" value={pinCur} onChange={(e) => { setPinCur(e.target.value); setPinMsg(null); }} autoComplete="new-password" /></label>
            <label className="fld"><span>New PIN (min 4 chars)</span><input className="inp" type="password" value={pinNew} onChange={(e) => { setPinNew(e.target.value); setPinMsg(null); }} autoComplete="new-password" /></label>
          </div>
          <button className="btn btn-primary" style={{ marginTop: 12, width: "auto" }} onClick={changeMyPin}>Update PIN</button>
          {pinMsg && <p className={"savedmsg " + pinMsg[0]}>{pinMsg[1]}</p>}
        </Card>
      </div>
    </>
  );
}

/* ================= PACKSHED ================= */
const ORDERS_KEY = "ml-orders-v1"; // shared: one team-wide order book
async function loadOrdersStore() {
  try { const r = await window.storage.get(ORDERS_KEY, true); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function saveOrdersStore(list) {
  try { await window.storage.set(ORDERS_KEY, JSON.stringify(list), true); return true; }
  catch { return false; }
}
const SIZE_OPTS = ["12MM+", "14MM+", "18MM+", "22MM+", "ALL"];
const sizeIdx = (sz) => ({ ALL: 0, "12MM+": 1, "14MM+": 2, "18MM+": 3, "22MM+": 4 })[sz] ?? 0;
const sum3 = (a) => n0(a[0]) + n0(a[1]) + n0(a[2]);
const DAYS = ["Sat", "Tue", "Thu"];

const STOCK_KEY = "ml-stock-v2"; // shared: live unprocessed-stock table (v2: starts empty, lines added manually)
async function loadStockStore() {
  try { const r = await window.storage.get(STOCK_KEY, true); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function saveStockStore(list) {
  try { await window.storage.set(STOCK_KEY, JSON.stringify(list), true); return true; }
  catch { return false; }
}
const GRADE_LABELS = ["12mm+", "14mm+", "18mm+", "22mm+", "2nd grade"];

function StockEditor({ init, isNew, onSave, onCancel, onDelete }) {
  const [x, setX] = useState({ ...init, gPct: init.g.map((v) => v * 100) });
  const setT = (k) => (e) => setX({ ...x, [k]: e.target.value });
  const setG = (i) => (e) => { const a = [...x.gPct]; a[i] = e.target.value; setX({ ...x, gPct: a }); };
  const setVariety = (e) => {
    const name = e.target.value;
    const preset = VARIETY_PRESETS[name];
    // selecting a variety loads its preset banding; the fields stay editable
    setX({ ...x, name, gPct: preset ? preset.map((v) => v * 100) : x.gPct });
  };
  const gSum = x.gPct.reduce((t, v) => t + (num(v) || 0), 0);
  const valid = String(x.name).trim();
  const save = () => {
    if (!valid) return;
    onSave({
      name: String(x.name).trim().toUpperCase(),
      stock: num(x.stock) || 0,
      g: x.gPct.map((v) => (num(v) || 0) / 100),
    });
  };
  return (
    <div className="oed">
      <div className="fgrid">
        <label className="fld"><span>Variety</span>
          <select className="inp" autoFocus value={x.name} onChange={setVariety}>
            <option value="">Select variety…</option>
            {STOCK_VARIETY_OPTS.map((v) => <option key={v}>{v}</option>)}
          </select>
        </label>
        <label className="fld"><span>In stock kg</span><input className="inp" type="number" step="any" value={x.stock} onChange={setT("stock")} /></label>
      </div>
      <div style={{ marginTop: 12 }}>
        <h4>Assumed grade-out (%) — preset loads on variety select, edit freely</h4>
        <div className="fgrid4">
          {GRADE_LABELS.map((l, i) => (
            <label key={l} className="fld"><span>{l}</span>
              <input className="inp" type="number" step="any" value={x.gPct[i]} onChange={setG(i)} /></label>
          ))}
        </div>
        <p className={"savedmsg " + (Math.abs(gSum - 100) < 0.5 ? "ok" : "warn")} style={{ marginTop: 8 }}>
          Grade-out total: {gSum.toFixed(1)}%{Math.abs(gSum - 100) >= 0.5 ? " — doesn't add to 100%" : ""}
        </p>
      </div>
      <div className="oed-actions">
        <button className="btn btn-primary" style={{ width: "auto" }} disabled={!valid} onClick={save}>
          {isNew ? "Add line" : "Save changes"}
        </button>
        <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
        {!isNew && <button className="btn btn-mini btn-danger" onClick={onDelete}>Remove line</button>}
      </div>
    </div>
  );
}

const EXTRA_KEY = "ml-otherstock-v1"; // shared: processed-not-packed + packed buffer, kg per size
const LOCAL_KEY = "ml-localsales-v1"; // shared: local sales stock lines
async function loadShared(key) {
  try { const r = await window.storage.get(key, true); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function saveShared(key, val) {
  try { await window.storage.set(key, JSON.stringify(val), true); return true; }
  catch { return false; }
}

function ExtraEditor({ label, init, onSave, onCancel }) {
  const [v, setV] = useState(init.map((x) => (x === 0 ? "" : x)));
  const setI = (i) => (e) => { const a = [...v]; a[i] = e.target.value; setV(a); };
  const total = v.reduce((t, x) => t + (num(x) || 0), 0);
  return (
    <div className="oed">
      <h4>{label} — kg per size</h4>
      <div className="fgrid4" style={{ marginTop: 8 }}>
        {PK_SIZES.map((sz, i) => (
          <label key={sz} className="fld"><span>{sz}</span>
            <input className="inp" type="number" step="any" value={v[i]} onChange={setI(i)} /></label>
        ))}
      </div>
      <p className="oed-kg">Total: {kg(total, 1)} kg</p>
      <div className="oed-actions">
        <button className="btn btn-primary" style={{ width: "auto" }} onClick={() => onSave(v.map((x) => num(x) || 0))}>Save changes</button>
        <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
      </div>
    </div>
  );
}

function LocalEditor({ init, isNew, onSave, onCancel, onDelete }) {
  const [x, setX] = useState(init);
  const setT = (k) => (e) => setX({ ...x, [k]: e.target.value });
  const unit = num(x.unit) || 0;
  const cartons = num(x.cartons) || 0;
  const valid = String(x.item).trim();
  return (
    <div className="oed">
      <div className="fgrid">
        <label className="fld"><span>Item</span><input className="inp" autoFocus value={x.item} onChange={setT("item")} /></label>
        <label className="fld"><span>Unit kg / carton</span><input className="inp" type="number" step="any" value={x.unit} onChange={setT("unit")} /></label>
        <label className="fld"><span>Cartons</span><input className="inp" type="number" step="any" value={x.cartons} onChange={setT("cartons")} /></label>
      </div>
      <p className="oed-kg">= {kg(cartons * unit, 1)} kg ({kg(cartons, 1)} ctns × {unit || "—"} kg)</p>
      <div className="oed-actions">
        <button className="btn btn-primary" style={{ width: "auto" }} disabled={!valid}
          onClick={() => onSave({ item: String(x.item).trim(), unit, cartons })}>
          {isNew ? "Add item" : "Save changes"}
        </button>
        <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
        {!isNew && <button className="btn btn-mini btn-danger" onClick={onDelete}>Remove item</button>}
      </div>
    </div>
  );
}

function OrderEditor({ init, isNew, onSave, onCancel, onDelete }) {
  const [o, setO] = useState(init);
  const setT = (k) => (e) => setO({ ...o, [k]: e.target.value });
  const setN = (k, i) => (e) => { const a = [...o[k]]; a[i] = e.target.value; setO({ ...o, [k]: a }); };
  const unit = num(o.unit) || 0;
  const ordCtn = sum3(o.ordCtn), pkdCtn = sum3(o.pkdCtn);
  const ordKg = ordCtn * unit, pkdKg = pkdCtn * unit;
  const valid = String(o.customer).trim() && o.size && num(o.week);
  const save = () => {
    if (!valid) return;
    const { ord, packed, ...rest } = o; // drop legacy kg arrays — cartons are the source of truth
    onSave({
      ...rest,
      week: Math.max(1, Math.min(53, Math.round(num(o.week) || 1))),
      customer: String(o.customer).trim().toUpperCase(),
      unit,
      ordCtn: o.ordCtn.map((x) => num(x) || 0),
      pkdCtn: o.pkdCtn.map((x) => num(x) || 0),
      status: o.status === "done" ? "done" : "prod",
      inc: [0, 1, 2].map((i2) => !!(o.inc || [])[i2]), // per-day: show on Week summary
    });
  };
  return (
    <div className="oed">
      <div className="fgrid">
        <label className="fld"><span>Week</span><input className="inp" type="number" min="1" max="53" value={o.week} onChange={setT("week")} /></label>
        <label className="fld"><span>Customer</span><input className="inp" autoFocus value={o.customer} onChange={setT("customer")} /></label>
        <label className="fld"><span>Size</span>
          <select className="inp" value={o.size} onChange={setT("size")}>
            {SIZE_OPTS.map((x) => <option key={x}>{x}</option>)}
          </select>
        </label>
        <label className="fld"><span>Unit weight (kg/carton)</span><input className="inp" type="number" step="any" value={o.unit} onChange={setT("unit")} /></label>
        <label className="fld"><span>Status</span>
          <select className="inp" value={o.status === "done" ? "done" : "prod"} onChange={setT("status")}>
            <option value="prod">In Production</option>
            <option value="done">Complete</option>
          </select>
        </label>
      </div>
      <div className="oed-days">
        <div>
          <h4>Cartons ordered</h4>
          <div className="fgrid3">
            {DAYS.map((d, i) => (
              <div key={d}>
                <label className="fld"><span>{d}</span>
                  <input className="inp" type="number" step="any" value={o.ordCtn[i]} onChange={setN("ordCtn", i)} /></label>
                <label style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 11.5, marginTop: 5, cursor: "pointer", color: "var(--muted)" }}>
                  <input type="checkbox" checked={!!(o.inc || [])[i]}
                    onChange={() => { const a = [0, 1, 2].map((x) => !!(o.inc || [])[x]); a[i] = !a[i]; setO({ ...o, inc: a }); }} />
                  Week summary
                </label>
              </div>
            ))}
          </div>
          <p className="oed-kg">= {kg(ordKg, 1)} kg ordered ({kg(ordCtn, 1)} ctns × {unit || "—"} kg)</p>
        </div>
        <div>
          <h4>Cartons packed</h4>
          <div className="fgrid3">
            {DAYS.map((d, i) => (
              <label key={d} className="fld"><span>{d}</span>
                <input className="inp" type="number" step="any" value={o.pkdCtn[i]} onChange={setN("pkdCtn", i)} /></label>
            ))}
          </div>
          <p className="oed-kg">= {kg(pkdKg, 1)} kg packed ({kg(pkdCtn, 1)} ctns × {unit || "—"} kg)</p>
        </div>
        <div className="oed-bal">
          <h4>Balance</h4>
          <div className={"oed-balv " + (ordKg - pkdKg > 0 ? "neg" : "pos")}>
            {kg(ordKg - pkdKg, 1)} kg
          </div>
          <p className="oed-kg">{kg(ordCtn - pkdCtn, 1)} cartons</p>
        </div>
      </div>
      <div className="oed-actions">
        <button className="btn btn-primary" style={{ width: "auto" }} disabled={!valid} onClick={save}>
          {isNew ? "Add order" : "Save changes"}
        </button>
        <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
        {!isNew && <button className="btn btn-mini btn-danger" onClick={onDelete}>Delete order</button>}
      </div>
    </div>
  );
}

function parseOrdersFile(data) {
  const wb = XLSX.read(data, { type: "array" });
  const isTxt = (c, t) => typeof c === "string" && c.trim().toUpperCase() === t;
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    const hi = rows.findIndex((r) => r && r.some((c) => isTxt(c, "CUSTOMER")) && r.some((c) => isTxt(c, "SIZE")));
    if (hi === -1) continue;
    const c0 = rows[hi].findIndex((c) => isTxt(c, "CUSTOMER"));
    const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const cust = r[c0];
      if (cust == null || String(cust).trim() === "") continue;
      if (String(cust).toUpperCase().includes("TOTAL")) continue;
      const unit = num(r[c0 + 2]) || 0;
      const g = (off) => [num(r[c0 + off]) || 0, num(r[c0 + off + 1]) || 0, num(r[c0 + off + 2]) || 0];
      let ordCtn = g(3);
      const ordKg = g(6);
      let pkdCtn = g(9);
      const pkdKg = g(12);
      // layouts that only give kgs: derive cartons from kg / unit
      if (!ordCtn.some(Boolean) && ordKg.some(Boolean) && unit) ordCtn = ordKg.map((v) => +(v / unit).toFixed(2));
      if (!pkdCtn.some(Boolean) && pkdKg.some(Boolean) && unit) pkdCtn = pkdKg.map((v) => +(v / unit).toFixed(2));
      const rawSize = r[c0 + 1] ? String(r[c0 + 1]).trim().toUpperCase() : "";
      const size = SIZE_OPTS.includes(rawSize) ? rawSize
        : rawSize.includes("12") ? "12MM+" : rawSize.includes("14") ? "14MM+"
        : rawSize.includes("18") ? "18MM+" : rawSize.includes("22") ? "22MM+" : "ALL";
      out.push({ customer: String(cust).trim().toUpperCase(), size, unit, ordCtn, pkdCtn });
    }
    if (out.length) return { sheet: name, rows: out };
  }
  return null;
}

function OrdersImport({ weekDefault, existingKeys, onImport }) {
  const [wk, setWk] = useState(weekDefault);
  useEffect(() => { setWk(weekDefault); }, [weekDefault]);
  const [parsed, setParsed] = useState(null);
  const [fileName, setFileName] = useState("");
  const [err, setErr] = useState("");
  const [done, setDone] = useState(null);
  const [busy, setBusy] = useState(false);
  const wkOk = !!num(wk) && num(wk) >= 1 && num(wk) <= 53;

  const onFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setErr(""); setDone(null); setParsed(null); setFileName(file.name); setBusy(true);
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      const res = parseOrdersFile(data);
      if (!res) setErr('No order lines found. The sheet needs a header row containing "CUSTOMER" and "SIZE" (like the Orders sheet in the packshed workbook).');
      else setParsed(res);
    } catch {
      setErr("Couldn't read this file. Supported: .xlsx, .xlsm, .csv");
    }
    setBusy(false);
  };

  const doImport = async () => {
    setBusy(true);
    const res = await onImport(Math.round(num(wk)), parsed.rows);
    setBusy(false);
    setParsed(null);
    setDone("Imported into week " + Math.round(num(wk)) + ": " + res.added + " added, " + res.updated + " updated." + (res.ok ? "" : " Shared storage is unavailable — changes may not persist."));
  };

  return (
    <Card title="Import orders from Excel" right={<span className="muted small">.xlsx · .xlsm · .csv · one file at a time</span>}>
      <p className="muted" style={{ marginBottom: 10 }}>
        Upload an Orders sheet and the lines are extracted automatically. Lines matching an existing order (same week, customer and size) update it; new lines are added.
      </p>
      <div className="fgrid" style={{ maxWidth: 200, marginBottom: 10 }}>
        <label className="fld"><span>Import into week</span>
          <input className="inp" type="number" min="1" max="53" value={wk} onChange={(e) => { setWk(e.target.value); setDone(null); }} />
        </label>
      </div>
      {!wkOk && <p className="savedmsg warn" style={{ marginBottom: 10 }}>Enter a valid week (1\u201353) before choosing a file.</p>}
      <input type="file" accept=".xlsx,.xlsm,.xls,.csv" onChange={onFile} className="filein" disabled={!wkOk || busy || !!parsed} />
      {busy && <p className="muted" style={{ marginTop: 10 }}>Reading…</p>}
      {err && <p className="savedmsg warn" style={{ marginTop: 10 }}>{err}</p>}
      {done && <p className="savedmsg ok" style={{ marginTop: 10 }}>{done}</p>}
      {parsed && (
        <div style={{ marginTop: 12 }}>
          <p style={{ marginBottom: 8 }}>
            <b>{fileName}</b> · sheet "{parsed.sheet}" · <b>{parsed.rows.length}</b> order line{parsed.rows.length === 1 ? "" : "s"}
          </p>
          <div className="tbl-wrap" style={{ maxHeight: 220, overflowY: "auto", border: "1px solid var(--line)", borderRadius: 8 }}>
            <table>
              <thead><tr><th>Customer</th><th>Size</th><th className="r">Unit kg</th><th className="r">Ctns ordered</th><th className="r">Ctns packed</th><th className="r">Balance kg</th><th></th></tr></thead>
              <tbody>
                {parsed.rows.map((o, i) => {
                  const upd = existingKeys.has(Math.round(num(wk)) + "|" + o.customer + "|" + o.size);
                  return (
                    <tr key={i}>
                      <td className="strong">{o.customer}</td><td>{o.size}</td>
                      <td className="r">{o.unit}</td>
                      <td className="r">{kg(sum3(o.ordCtn), 1)}</td>
                      <td className="r">{kg(sum3(o.pkdCtn), 1)}</td>
                      <td className="r">{kg((sum3(o.ordCtn) - sum3(o.pkdCtn)) * o.unit, 1)}</td>
                      <td className="muted small">{upd ? "updates existing" : "new"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="oed-actions">
            <button className="btn btn-primary" style={{ width: "auto" }} disabled={busy} onClick={doImport}>
              Import {parsed.rows.length} line{parsed.rows.length === 1 ? "" : "s"}
            </button>
            <button className="btn btn-mini" onClick={() => setParsed(null)}>Cancel</button>
          </div>
        </div>
      )}
    </Card>
  );
}

function Packshed({ me }) {
  const P = PACKSHED;
  const canEdit = me.role === "admin" || me.role === "entry";
  const [orders, setOrders] = useState(null);
  const [editing, setEditing] = useState(null); // order id, "new", or null
  const [storeWarn, setStoreWarn] = useState(false);
  const [week, setWeek] = useState(String(P.week));
  const [tab, setTab] = useState("orders");
  const [day, setDay] = useState(""); // "" = all days, else Saturday/Tuesday/Thursday

  const [stock, setStock] = useState(null);
  const [editStock, setEditStock] = useState(null); // stock id, "new", or null
  const [extra, setExtra] = useState(null); // { proc: [5], buffer: [5] } kg per size
  const [editExtra, setEditExtra] = useState(null); // "proc" | "buffer" | null
  const [localSales, setLocalSales] = useState(null);
  const [editLocal, setEditLocal] = useState(null); // local id, "new", or null

  const migrateOrders = (list) => {
    // cartons are the source of truth; older records stored kg — derive ctns = kg / unit
    const toCtn = (kgArr, unit) => (kgArr || [0, 0, 0]).map((v) => (unit ? +(n0(v) / unit).toFixed(2) : 0));
    return list.map((o) => ({
      ...o,
      week: o.week ?? P.week,
      ordCtn: o.ordCtn ?? toCtn(o.ord, n0(o.unit)),
      pkdCtn: o.pkdCtn ?? toCtn(o.packed, n0(o.unit)),
    }));
  };
  const migrateLocal = (list) => {
    // kg is cartons x unit; derive unit from the item name ("5 kg" -> 5), else from kg/cartons
    const unitOf = (x) => {
      if (x.unit != null) return n0(x.unit);
      const m = String(x.item).match(/(\d+(?:\.\d+)?)\s*kg/i);
      if (m) return +m[1];
      return n0(x.cartons) ? +(n0(x.kgs) / n0(x.cartons)).toFixed(2) : 0;
    };
    return list.map((x) => ({ id: x.id, item: x.item, cartons: n0(x.cartons), unit: unitOf(x) }));
  };

  useEffect(() => {
    loadOrdersStore().then((saved) => setOrders(migrateOrders(saved || P.orders.map((o, i) => ({ ...o, id: "wb-" + i })))));
    loadStockStore().then((saved) => setStock(saved || [])); // starts empty — lines are added manually
    loadShared(EXTRA_KEY).then((saved) => setExtra(saved || { proc: P.procNotPacked.slice(0, 5), buffer: P.packedBuffer.slice(0, 5) }));
    loadShared(LOCAL_KEY).then((saved) => setLocalSales(migrateLocal(saved || P.localSales.map((x, i) => ({ ...x, id: "ls-" + i })))));
  }, []);

  // Live refresh: pick up teammates' packshed changes without restarting the app.
  useEffect(() => {
    const typing = () => {
      const el = document.activeElement;
      return !!el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.tagName === "SELECT");
    };
    const refresh = async () => {
      if (typing()) return;
      try {
        const [o2, s2, e2, l2] = await Promise.all([
          loadOrdersStore(), loadStockStore(), loadShared(EXTRA_KEY), loadShared(LOCAL_KEY),
        ]);
        if (o2) { const m2 = migrateOrders(o2); setOrders((cur) => (JSON.stringify(m2) !== JSON.stringify(cur) ? m2 : cur)); }
        if (s2) setStock((cur) => (JSON.stringify(s2) !== JSON.stringify(cur) ? s2 : cur));
        if (e2) setExtra((cur) => (JSON.stringify(e2) !== JSON.stringify(cur) ? e2 : cur));
        if (l2) { const m3 = migrateLocal(l2); setLocalSales((cur) => (JSON.stringify(m3) !== JSON.stringify(cur) ? m3 : cur)); }
      } catch (e) { /* transient — retry next tick */ }
    };
    const t = setInterval(refresh, 20000);
    const onWake = () => refresh();
    window.addEventListener("focus", onWake);
    document.addEventListener("visibilitychange", onWake);
    return () => {
      clearInterval(t);
      window.removeEventListener("focus", onWake);
      document.removeEventListener("visibilitychange", onWake);
    };
  }, []);

  const all = orders || [];
  const weeks = [...new Set(all.map((o) => o.week))].sort((a, b) => b - a); // newest first
  const live = all.filter((o) => !week || o.week === +week);
  const persist = async (list) => {
    setOrders(list); setEditing(null);
    setStoreWarn(!(await saveOrdersStore(list)));
  };
  const persistStock = async (list) => {
    setStock(list); setEditStock(null);
    setStoreWarn(!(await saveStockStore(list)));
  };
  const persistExtra = async (next) => {
    setExtra(next); setEditExtra(null);
    setStoreWarn(!(await saveShared(EXTRA_KEY, next)));
  };
  const persistLocal = async (list) => {
    setLocalSales(list); setEditLocal(null);
    setStoreWarn(!(await saveShared(LOCAL_KEY, list)));
  };
  const liveExtra = extra || { proc: P.procNotPacked.slice(0, 5), buffer: P.packedBuffer.slice(0, 5) };
  const liveLocal = localSales || [];
  const liveStock = stock || [];
  // projections: g = [12,14,18,22,2nd]; "ALL" bucket = 2nd-grade kg
  const projOf = (x) => [x.stock * x.g[4], x.stock * x.g[0], x.stock * x.g[1], x.stock * x.g[2], x.stock * x.g[3]];
  const unprocBy = [0, 0, 0, 0, 0];
  liveStock.forEach((x) => projOf(x).forEach((v, i) => (unprocBy[i] += v)));
  const availBy = unprocBy.map((v, i) => v + n0(liveExtra.proc[i]) + n0(liveExtra.buffer[i]));
  const importOrders = async (wk, rows) => {
    let added = 0, updated = 0;
    const next = [...all];
    const stamp = Date.now();
    for (const r of rows) {
      const i = next.findIndex((o) => o.week === wk && o.customer === r.customer && o.size === r.size);
      const rec = { ...r, week: wk };
      if (i >= 0) { next[i] = { ...rec, id: next[i].id, status: next[i].status === "done" ? "done" : "prod", inc: next[i].inc || [false, false, false] }; updated++; }
      else { next.push({ ...rec, id: "o-" + stamp + "-" + (added + updated), status: "prod", inc: [false, false, false] }); added++; }
    }
    setOrders(next); setEditing(null);
    const ok = await saveOrdersStore(next);
    setStoreWarn(!ok);
    if (week && week !== String(wk)) setWeek(String(wk)); // jump to the imported week
    return { added, updated, ok };
  };
  const orderKeys = useMemo(() => new Set(all.map((o) => o.week + "|" + o.customer + "|" + o.size)), [all]);
  const addOrder = (o) => {
    const rec = { ...o, id: "o-" + Date.now() };
    persist([...all, rec]);
    if (week && rec.week !== +week) setWeek(String(rec.week)); // jump to the week you just added to
  };

  // live demand aggregates by size (ALL,12,14,18,22)
  const di = day ? { Saturday: 0, Tuesday: 1, Thursday: 2 }[day] : null;
  const ctnOf = (o, k) => (di == null ? sum3(o[k]) : n0(o[k][di]));
  const kgOf = (o, k) => ctnOf(o, k) * n0(o.unit);
  const shown = di == null ? live : live.filter((o) => ctnOf(o, "ordCtn") > 0 || ctnOf(o, "pkdCtn") > 0);
  const ordBy = [0, 0, 0, 0, 0], pkdBy = [0, 0, 0, 0, 0];
  live.forEach((o) => { const i = sizeIdx(o.size); ordBy[i] += kgOf(o, "ordCtn"); pkdBy[i] += kgOf(o, "pkdCtn"); });
  const totOrdCtn = live.reduce((t, o) => t + ctnOf(o, "ordCtn"), 0);
  const totPkdCtn = live.reduce((t, o) => t + ctnOf(o, "pkdCtn"), 0);
  const balBy = ordBy.map((x, i) => x - pkdBy[i]);
  const surplusBy = balBy.map((b, i) => availBy[i] - b);
  const totOrd = ordBy.reduce((a, b) => a + b, 0);
  const totPacked = pkdBy.reduce((a, b) => a + b, 0);
  const totBal = totOrd - totPacked;
  const totAvail = availBy.reduce((a, b) => a + b, 0);
  const totSurplus = totAvail - totBal;

  const surplusChart = PK_SIZES.map((sz, i) => ({ name: sz, kg: +surplusBy[i].toFixed(0) }));

  // Packshed-update summary matrix (whole week, day filter not applied)
  const sumOrders = live.filter((o) => o.status !== "done"); // Complete orders drop off the Week summary
  const dayKg = (key, d) => {
    const a = [0, 0, 0, 0, 0];
    sumOrders.forEach((o) => {
      if (!(o.inc && o.inc[d])) return; // only days ticked "Week summary" in the order editor
      a[sizeIdx(o.size)] += n0(o[key][d]) * n0(o.unit);
    });
    return a;
  };
  const ordD = [0, 1, 2].map((d) => dayKg("ordCtn", d));
  const pkdD = [0, 1, 2].map((d) => dayKg("pkdCtn", d));
  const balD = [0, 1, 2].map((d) => ordD[d].map((v, i) => v - pkdD[d][i]));
  const colSum = (rows2) => [0, 1, 2, 3, 4].map((i) => rows2.reduce((t, r) => t + r[i], 0));
  const ordT = colSum(ordD), pkdT = colSum(pkdD), balT = colSum(balD);
  const supplyRow = unprocBy.map((v, i) => v + n0(liveExtra.proc[i])); // Excel: unprocessed + processed-not-packed
  const surplusRow = supplyRow.map((v, i) => v - balT[i]);
  const SUM_DAYS = ["SAT/SUN", "TUES", "THUR"];
  const SRow = ({ label, vals, strong, redNeg, redZero }) => (
    <tr className={strong ? "totalrow" : ""}>
      <td className={strong ? "strong" : ""}>{label}</td>
      {vals.map((v, i) => (
        <td key={i} className={"r " + (strong ? "strong " : "") + ((redNeg && v < 0) || (redZero && !n0(v)) ? "neg" : "")}>{kg(v, 2)}</td>
      ))}
      <td className={"r strong " + (redNeg && vals.reduce((a, b) => a + b, 0) < 0 ? "neg" : "")}>{kg(vals.reduce((a, b) => a + b, 0), 2)}</td>
    </tr>
  );
  const localTot = liveLocal.reduce((t, x) => ({ cartons: t.cartons + n0(x.cartons), kgs: t.kgs + n0(x.cartons) * n0(x.unit) }), { cartons: 0, kgs: 0 });
  const blankOrder = { week: week ? +week : isoWeek(new Date().toISOString().slice(0, 10)), customer: "", size: "18MM+", unit: 2, ordCtn: ["", "", ""], pkdCtn: ["", "", ""] };

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Packshed update</h2>
          <p className="muted">{week ? "Week " + week : "All weeks"}{day ? " · " + day + " only" : ""} · stock snapshot wk {P.week}, {new Date(P.asOf).toLocaleString("en-GB", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })} · orders are live and shared with everyone using this app</p>
        </div>
        <div className="controls">
          <label className="lbl">Week</label>
          <Select value={week} onChange={setWeek} options={weeks.map(String)} all="All weeks" />
          <label className="lbl">Day</label>
          <Select value={day} onChange={setDay} options={["Saturday", "Tuesday", "Thursday"]} all="All days" />
          <PdfBtn name="Packshed" />
        </div>
      </div>

      <div className="kpis">
        <Kpi label="Orders (wk)" value={kg(totOrd) + " kg"} sub="all customers, Sat + Tue + Thu" />
        <Kpi label="Packed" value={kg(totPacked) + " kg"} sub={totOrd ? pc(totPacked / totOrd, 0) + " of orders" : " "} />
        <Kpi label="Balance to pack" value={kg(totBal) + " kg"} tone={totBal > 0 ? "bad" : "good"} sub="remaining on this week's orders" />
        <Kpi label="Est. available" value={kg(totAvail) + " kg"} sub="unpacked + unprocessed at assumed grade-out" />
        <Kpi label="Projected surplus" value={(totSurplus >= 0 ? "+" : "") + kg(totSurplus) + " kg"} tone={totSurplus >= 0 ? "good" : "bad"} sub="available minus balance to pack" />
      </div>

      <div className="tabs">
        <button className={"tab" + (tab === "orders" ? " active" : "")} onClick={() => setTab("orders")}>Customer orders</button>
        <button className={"tab" + (tab === "update" ? " active" : "")} onClick={() => setTab("update")}>Packshed update</button>
        <button className={"tab" + (tab === "summary" ? " active" : "")} onClick={() => setTab("summary")}>Week summary</button>
      </div>

      {tab === "orders" && (<>
      <Card
        title={week ? "Customer orders — week " + week : "Customer orders — all weeks"}
        right={canEdit
          ? <button className="btn btn-primary" style={{ width: "auto", padding: "6px 12px" }} onClick={() => setEditing("new")}>+ Add order</button>
          : <span className="muted small">view only</span>}
        pad={false}
      >
        {editing === "new" && (
          <OrderEditor init={blankOrder} isNew
            onSave={addOrder}
            onCancel={() => setEditing(null)} />
        )}
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>{!week && <th className="r">Wk</th>}<th>Customer</th><th>Size</th><th className="r">Unit kg</th>
                <th className="r">Ctns ordered</th><th className="r">Ordered kg</th>
                <th className="r">Ctns packed</th><th className="r">Packed kg</th>
                <th className="r">Balance kg</th>
                {canEdit && <th className="r"></th>}</tr>
            </thead>
            <tbody>
              {shown.map((o) => {
                const oKg = kgOf(o, "ordCtn"), pKg = kgOf(o, "pkdCtn");
                const bal = oKg - pKg;
                const isEd = editing === o.id;
                return (
                  <React.Fragment key={o.id}>
                    <tr className={bal > 0 ? "hidef" : ""}>
                      {!week && <td className="r muted">{o.week}</td>}
                      <td className="strong">{o.customer}{o.status === "done" && <span className="msg-role" style={{ marginLeft: 6 }}>COMPLETE</span>}</td>
                      <td>{o.size}</td>
                      <td className="r">{o.unit}</td>
                      <td className="r">{kg(ctnOf(o, "ordCtn"), 1)}</td>
                      <td className="r">{kg(oKg, 1)}</td>
                      <td className="r">{kg(ctnOf(o, "pkdCtn"), 1)}</td>
                      <td className="r">{kg(pKg, 1)}</td>
                      <td className={"r " + (bal > 0 ? "neg strong" : "")}>{kg(bal, 1)}</td>
                      {canEdit && (
                        <td className="r">
                          <button className="btn btn-mini" onClick={() => setEditing(isEd ? null : o.id)}>{isEd ? "Close" : "Edit"}</button>
                        </td>
                      )}
                    </tr>
                    {isEd && (
                      <tr className="detail"><td colSpan={week ? 9 : 10}>
                        <OrderEditor init={o}
                          onSave={(n) => persist(all.map((x) => (x.id === o.id ? { ...n, id: o.id } : x)))}
                          onCancel={() => setEditing(null)}
                          onDelete={() => persist(all.filter((x) => x.id !== o.id))} />
                      </td></tr>
                    )}
                  </React.Fragment>
                );
              })}
              <tr>
                {!week && <td></td>}
                <td className="strong">TOTAL</td><td></td><td></td>
                <td className="r strong">{kg(totOrdCtn, 1)}</td>
                <td className="r strong">{kg(totOrd, 1)}</td>
                <td className="r strong">{kg(totPkdCtn, 1)}</td>
                <td className="r strong">{kg(totPacked, 1)}</td>
                <td className="r strong">{kg(totBal, 1)}</td>
                {canEdit && <td></td>}
              </tr>
            </tbody>
          </table>
        </div>
        {storeWarn && <p className="savedmsg warn" style={{ padding: "0 16px 14px" }}>Changes are shown for this session, but shared storage is unavailable — they may not persist.</p>}
        {orders === null && <p className="muted" style={{ padding: "0 16px 14px" }}>Loading saved orders…</p>}
      </Card>

      {canEdit && (
        <OrdersImport
          weekDefault={week || String(isoWeek(new Date().toISOString().slice(0, 10)))}
          existingKeys={orderKeys}
          onImport={importOrders}
        />
      )}
      </>
      )}

      {tab === "summary" && (<>
      <Card title={"Packshed update summary" + (week ? " \u2014 week " + week : "")}
        right={<span className="muted small">whole week {"\u00b7"} day filter not applied</span>} pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr><th style={{ minWidth: 190 }}></th>{PK_SIZES.map((sz) => <th key={sz} className="r">{sz.toUpperCase()}</th>)}<th className="r">Total</th></tr>
            </thead>
            <tbody>
              <tr className="dayhead"><td colSpan={7}>Unprocessed product (assumed grade-out %)</td></tr>
              {liveStock.map((x) => {
                const pr = projOf(x);
                return (
                  <tr key={x.id}>
                    <td>
                      <span className="strong">{x.name}</span>
                      <div className="muted small">{"12: " + Math.round(x.g[0] * 100) + "% \u00b7 14: " + Math.round(x.g[1] * 100) + "% \u00b7 18: " + Math.round(x.g[2] * 100) + "% \u00b7 22: " + Math.round(x.g[3] * 100) + "% \u00b7 2nd: " + Math.round(x.g[4] * 100) + "%"}</div>
                    </td>
                    {pr.map((v, i) => <td key={i} className="r">{v ? kg(v, 2) : "\u2014"}</td>)}
                    <td className="r strong">{kg(n0(x.stock), 2)}</td>
                  </tr>
                );
              })}
              {!liveStock.length && <tr><td colSpan={7} className="muted center">No unprocessed stock entered (Packshed update tab).</td></tr>}
              <SRow label="TOTAL unprocessed" vals={unprocBy} strong />
              <SRow label="Processed not packed" vals={liveExtra.proc.map(n0)} redZero />
              <SRow label="Packed buffer stock" vals={liveExtra.buffer.map(n0)} redZero />

              <tr className="dayhead"><td colSpan={7}>Ordered kgs — only days ticked "Week summary" in the order editor</td></tr>
              {SUM_DAYS.map((dl, d) => <SRow key={dl} label={dl} vals={ordD[d]} />)}
              <SRow label="TOTAL orders kgs" vals={ordT} strong />

              <tr className="dayhead"><td colSpan={7}>Ordered kgs packed</td></tr>
              {SUM_DAYS.map((dl, d) => <SRow key={dl} label={dl} vals={pkdD[d]} />)}
              <SRow label="TOTAL packed" vals={pkdT} strong />

              <tr className="dayhead"><td colSpan={7}>Balance to be packed</td></tr>
              {SUM_DAYS.map((dl, d) => <SRow key={dl} label={dl} vals={balD[d]} />)}
              <SRow label="TOTAL balance" vals={balT} strong />

              <tr className="dayhead"><td colSpan={7}>Position</td></tr>
              <SRow label="Processed not packed + unprocessed" vals={supplyRow} strong />
              <SRow label="Surplus / shortfall vs balance" vals={surplusRow} strong redNeg />
            </tbody>
          </table>
        </div>
      </Card>

      <Card title="Local sales stock" pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead><tr><th>Item</th><th className="r">Cartons</th><th className="r">Kgs</th></tr></thead>
            <tbody>
              {liveLocal.map((x) => (
                <tr key={x.id}>
                  <td className="strong">{x.item}</td>
                  <td className="r">{kg(n0(x.cartons), 0)}</td>
                  <td className="r">{kg(x.kgs != null ? n0(x.kgs) : n0(x.cartons) * n0(x.unit), 2)}</td>
                </tr>
              ))}
              {!liveLocal.length && <tr><td colSpan={3} className="muted center">No local sales lines.</td></tr>}
              <tr className="totalrow">
                <td className="strong">TOTAL</td>
                <td className="r strong">{kg(liveLocal.reduce((t, x) => t + n0(x.cartons), 0), 0)}</td>
                <td className="r strong">{kg(liveLocal.reduce((t, x) => t + (x.kgs != null ? n0(x.kgs) : n0(x.cartons) * n0(x.unit)), 0), 2)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>
      </>)}

      {tab === "update" && (<>
      <Card title="Orders by pack day" right={<span className="muted small">Saturday \u2192 Tuesday \u2192 Thursday</span>} pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr><th>Customer</th><th>Size</th><th className="r">Unit kg</th>
                <th className="r">Ctns ordered</th><th className="r">Ordered kg</th>
                <th className="r">Ctns packed</th><th className="r">Packed kg</th>
                <th className="r">Balance kg</th></tr>
            </thead>
            <tbody>
              {[["Saturday", 0], ["Tuesday", 1], ["Thursday", 2]].filter(([l]) => !day || l === day).map(([label, d]) => {
                const lines = live.filter((o) => n0(o.ordCtn[d]) > 0 || n0(o.pkdCtn[d]) > 0);
                const tOrdC = lines.reduce((t, o) => t + n0(o.ordCtn[d]), 0);
                const tPkdC = lines.reduce((t, o) => t + n0(o.pkdCtn[d]), 0);
                const tOrdK = lines.reduce((t, o) => t + n0(o.ordCtn[d]) * n0(o.unit), 0);
                const tPkdK = lines.reduce((t, o) => t + n0(o.pkdCtn[d]) * n0(o.unit), 0);
                return (
                  <React.Fragment key={label}>
                    <tr className="dayhead"><td colSpan={8}>{label}</td></tr>
                    {lines.map((o) => {
                      const bal = (n0(o.ordCtn[d]) - n0(o.pkdCtn[d])) * n0(o.unit);
                      return (
                        <tr key={o.id + "-" + d} className={bal > 0 ? "hidef" : ""}>
                          <td className="strong">{o.customer}</td><td>{o.size}</td>
                          <td className="r">{o.unit}</td>
                          <td className="r">{kg(n0(o.ordCtn[d]), 1)}</td>
                          <td className="r">{kg(n0(o.ordCtn[d]) * n0(o.unit), 1)}</td>
                          <td className="r">{kg(n0(o.pkdCtn[d]), 1)}</td>
                          <td className="r">{kg(n0(o.pkdCtn[d]) * n0(o.unit), 1)}</td>
                          <td className={"r " + (bal > 0 ? "neg strong" : "")}>{kg(bal, 1)}</td>
                        </tr>
                      );
                    })}
                    {!lines.length && <tr><td colSpan={8} className="muted center">No orders for {label}.</td></tr>}
                    <tr className="totalrow">
                      <td className="strong" colSpan={3}>{label} total</td>
                      <td className="r strong">{kg(tOrdC, 1)}</td><td className="r strong">{kg(tOrdK, 1)}</td>
                      <td className="r strong">{kg(tPkdC, 1)}</td><td className="r strong">{kg(tPkdK, 1)}</td>
                      <td className={"r strong " + (tOrdK - tPkdK > 0 ? "neg" : "")}>{kg(tOrdK - tPkdK, 1)}</td>
                    </tr>
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      <div className="grid2">
        <Card title="Supply vs demand by size" right={<span className="muted small">demand from {week ? "week " + week : "all"} orders · availability live from stock table + wk {P.week} buffers</span>} pad={false}>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Size</th><th className="r">Ordered kg</th><th className="r">Packed kg</th>
                  <th className="r">Balance</th><th className="r">Est. available</th><th className="r">Surplus / shortfall</th></tr>
              </thead>
              <tbody>
                {PK_SIZES.map((sz, i) => (
                  <tr key={sz} className={surplusBy[i] < 0 ? "hidef" : ""}>
                    <td className="strong">{sz}</td>
                    <td className="r">{kg(ordBy[i])}</td>
                    <td className="r">{kg(pkdBy[i])}</td>
                    <td className="r">{kg(balBy[i])}</td>
                    <td className="r">{kg(availBy[i])}</td>
                    <td className={"r strong " + (surplusBy[i] < 0 ? "neg" : "pos")}>
                      {(surplusBy[i] >= 0 ? "+" : "") + kg(surplusBy[i])}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>

        <Card title="Projected surplus / shortfall" right={<span className="muted small">kg by size</span>}>
          <SBarChart data={surplusChart} xKey="name" height={230} fmt={(v) => kg(v)}
            series={[{ key: "kg", label: "Surplus", colorFor: (v) => (v < 0 ? "#AE4038" : "#2F7D53"), color: "#2F7D53" }]} />
        </Card>
      </div>

      <Card
        title="Unprocessed stock by variety"
        right={canEdit
          ? <button className="btn btn-primary" style={{ width: "auto", padding: "6px 12px" }} onClick={() => setEditStock("new")}>+ Add line</button>
          : <span className="muted small">view only</span>}
        pad={false}
      >
        {editStock === "new" && (
          <StockEditor init={{ name: "", stock: "", g: [0, 0, 0, 0, 0] }} isNew
            onSave={(x) => persistStock([...liveStock, { ...x, id: "st-" + Date.now() }])}
            onCancel={() => setEditStock(null)} />
        )}
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr><th>Variety</th><th className="r">In stock kg</th>
                <th className="r">12mm+ %</th><th className="r">14mm+ %</th><th className="r">18mm+ %</th><th className="r">22mm+ %</th><th className="r">2nd grade %</th>
                <th className="r">→ 12mm+</th><th className="r">→ 14mm+</th><th className="r">→ 18mm+</th><th className="r">→ 22mm+</th>
                {canEdit && <th className="r"></th>}</tr>
            </thead>
            <tbody>
              {liveStock.map((x) => {
                const pr = projOf(x); // [ALL,12,14,18,22]
                const isEd = editStock === x.id;
                return (
                  <React.Fragment key={x.id}>
                    <tr>
                      <td className="strong">{x.name}</td>
                      <td className="r">{kg(x.stock)}</td>
                      {x.g.map((g, i) => <td key={i} className="r">{pc(g, 0)}</td>)}
                      <td className="r">{kg(pr[1])}</td><td className="r">{kg(pr[2])}</td>
                      <td className="r">{kg(pr[3])}</td><td className="r">{kg(pr[4])}</td>
                      {canEdit && (
                        <td className="r">
                          <button className="btn btn-mini" onClick={() => setEditStock(isEd ? null : x.id)}>{isEd ? "Close" : "Edit"}</button>
                        </td>
                      )}
                    </tr>
                    {isEd && (
                      <tr className="detail"><td colSpan={canEdit ? 12 : 11}>
                        <StockEditor init={x}
                          onSave={(n) => persistStock(liveStock.map((y) => (y.id === x.id ? { ...n, id: x.id } : y)))}
                          onCancel={() => setEditStock(null)}
                          onDelete={() => persistStock(liveStock.filter((y) => y.id !== x.id))} />
                      </td></tr>
                    )}
                  </React.Fragment>
                );
              })}
              {!liveStock.length && <tr><td colSpan={canEdit ? 12 : 11} className="muted center">No stock lines. {canEdit ? "Use + Add line to create one." : ""}</td></tr>}
              <tr>
                <td className="strong">TOTAL</td>
                <td className="r strong">{kg(liveStock.reduce((t, x) => t + n0(x.stock), 0))}</td>
                <td colSpan={5}></td>
                <td className="r strong">{kg(unprocBy[1])}</td><td className="r strong">{kg(unprocBy[2])}</td>
                <td className="r strong">{kg(unprocBy[3])}</td><td className="r strong">{kg(unprocBy[4])}</td>
                {canEdit && <td></td>}
              </tr>
            </tbody>
          </table>
        </div>
        {stock === null && <p className="muted" style={{ padding: "0 16px 14px" }}>Loading saved stock lines…</p>}
      </Card>

      <div className="grid2">
        <Card title="Other stock positions" right={canEdit ? <span className="muted small">counted toward availability</span> : <span className="muted small">view only</span>} pad={false}>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Position</th><th className="r">ALL</th><th className="r">12mm+</th><th className="r">14mm+</th><th className="r">18mm+</th><th className="r">22mm+</th><th className="r">Total kg</th>
                  {canEdit && <th className="r"></th>}</tr>
              </thead>
              <tbody>
                {[["proc", "Processed, not packed"], ["buffer", "Packed buffer stock"]].map(([key, label]) => {
                  const row = liveExtra[key];
                  const tot = row.reduce((t, x) => t + n0(x), 0);
                  const isEd = editExtra === key;
                  return (
                    <React.Fragment key={key}>
                      <tr>
                        <td className="strong">{label}</td>
                        {row.map((v, i) => <td key={i} className="r">{kg(v)}</td>)}
                        <td className="r strong">{kg(tot)}</td>
                        {canEdit && (
                          <td className="r">
                            <button className="btn btn-mini" onClick={() => setEditExtra(isEd ? null : key)}>{isEd ? "Close" : "Edit"}</button>
                          </td>
                        )}
                      </tr>
                      {isEd && (
                        <tr className="detail"><td colSpan={canEdit ? 8 : 7}>
                          <ExtraEditor label={label} init={row}
                            onSave={(v) => persistExtra({ ...liveExtra, [key]: v })}
                            onCancel={() => setEditExtra(null)} />
                        </td></tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
          {extra === null && <p className="muted" style={{ padding: "0 16px 14px" }}>Loading…</p>}
        </Card>

        <Card
          title="Local sales stock"
          right={canEdit
            ? <button className="btn btn-primary" style={{ width: "auto", padding: "6px 12px" }} onClick={() => setEditLocal("new")}>+ Add item</button>
            : <span className="muted small">view only</span>}
          pad={false}
        >
          {editLocal === "new" && (
            <LocalEditor init={{ item: "", unit: "", cartons: "" }} isNew
              onSave={(x) => persistLocal([...liveLocal, { ...x, id: "ls-" + Date.now() }])}
              onCancel={() => setEditLocal(null)} />
          )}
          <div className="tbl-wrap">
            <table>
              <thead><tr><th>Item</th><th className="r">Unit kg</th><th className="r">Cartons</th><th className="r">Kg</th>{canEdit && <th className="r"></th>}</tr></thead>
              <tbody>
                {liveLocal.map((x) => {
                  const isEd = editLocal === x.id;
                  return (
                    <React.Fragment key={x.id}>
                      <tr>
                        <td className="strong">{x.item}</td>
                        <td className="r">{x.unit}</td>
                        <td className="r">{kg(x.cartons)}</td>
                        <td className="r">{kg(n0(x.cartons) * n0(x.unit), 1)}</td>
                        {canEdit && (
                          <td className="r">
                            <button className="btn btn-mini" onClick={() => setEditLocal(isEd ? null : x.id)}>{isEd ? "Close" : "Edit"}</button>
                          </td>
                        )}
                      </tr>
                      {isEd && (
                        <tr className="detail"><td colSpan={canEdit ? 5 : 4}>
                          <LocalEditor init={x}
                            onSave={(n) => persistLocal(liveLocal.map((y) => (y.id === x.id ? { ...n, id: x.id } : y)))}
                            onCancel={() => setEditLocal(null)}
                            onDelete={() => persistLocal(liveLocal.filter((y) => y.id !== x.id))} />
                        </td></tr>
                      )}
                    </React.Fragment>
                  );
                })}
                {!liveLocal.length && <tr><td colSpan={canEdit ? 5 : 4} className="muted center">No local sales items.</td></tr>}
                <tr><td className="strong">TOTAL</td><td></td>
                  <td className="r strong">{kg(localTot.cartons)}</td><td className="r strong">{kg(localTot.kgs, 1)}</td>
                  {canEdit && <td></td>}</tr>
              </tbody>
            </table>
          </div>
          {localSales === null && <p className="muted" style={{ padding: "0 16px 14px" }}>Loading…</p>}
        </Card>
      </div>
      </>)}
    </>
  );
}

/* ================= TEAM CHAT ================= */
const CHAT_KEY = "ml-chat-v1"; // shared: team messages (last 200 kept)

function Chat({ me }) {
  const [msgs, setMsgs] = useState(null);
  const [text, setText] = useState("");
  const [sending, setSending] = useState(false);
  const [warn, setWarn] = useState(false);
  const endRef = React.useRef(null);
  const countRef = React.useRef(0);

  const load = async () => {
    const list = (await loadShared(CHAT_KEY)) || [];
    setMsgs(list);
    return list;
  };

  useEffect(() => {
    load();
    const t = setInterval(load, 5000); // pick up teammates' messages
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    if (msgs && msgs.length !== countRef.current) {
      countRef.current = msgs.length;
      if (endRef.current) endRef.current.scrollIntoView({ block: "end" });
    }
  }, [msgs]);

  const send = async () => {
    const body = text.trim();
    if (!body || sending) return;
    setSending(true);
    const msg = {
      id: "m-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
      user: me.username, name: me.name, role: me.role,
      text: body.slice(0, 2000), ts: Date.now(),
    };
    const fresh = (await loadShared(CHAT_KEY)) || []; // re-read to avoid clobbering others
    const next = [...fresh, msg].slice(-200);
    const ok = await saveShared(CHAT_KEY, next);
    setWarn(!ok);
    setMsgs(next);
    setText("");
    setSending(false);
  };

  const del = async (id) => {
    const fresh = (await loadShared(CHAT_KEY)) || [];
    const next = fresh.filter((m) => m.id !== id);
    const ok = await saveShared(CHAT_KEY, next);
    setWarn(!ok);
    setMsgs(next);
  };

  const tfmt = (ts) => new Date(ts).toLocaleString("en-GB", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" });

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Team chat</h2>
          <p className="muted">Visible to everyone signed in · refreshes every few seconds · last 200 messages kept</p>
        </div>
      </div>

      <div className="chat">
        <div className="chat-msgs">
          {msgs === null && <p className="muted center">Loading messages…</p>}
          {msgs && !msgs.length && <p className="muted center">No messages yet — say hello.</p>}
          {msgs && msgs.map((m) => {
            const mine = m.user === me.username;
            const canDel = mine || me.role === "admin";
            return (
              <div key={m.id} className={"msg" + (mine ? " mine" : "")}>
                <div className="msg-meta">
                  <b>{mine ? "You" : m.name || m.user}</b>
                  {!mine && m.role && <span className="msg-role">{ROLE_LABEL[m.role] || m.role}</span>}
                  <span>{tfmt(m.ts)}</span>
                  {canDel && <button className="msg-del" title="Delete message" onClick={() => del(m.id)}>×</button>}
                </div>
                <div className="msg-bubble">{m.text}</div>
              </div>
            );
          })}
          <div ref={endRef} />
        </div>
        <div className="chat-input">
          <input
            className="inp" style={{ flex: 1 }} placeholder="Message the team…"
            value={text} maxLength={2000}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && send()}
          />
          <button className="btn btn-primary" style={{ width: "auto" }} disabled={!text.trim() || sending} onClick={send}>
            {sending ? "Sending…" : "Send"}
          </button>
        </div>
        {warn && <p className="savedmsg warn" style={{ marginTop: 8 }}>Shared storage is unavailable — messages may not reach the team.</p>}
      </div>
    </>
  );
}

/* --- Signature Pad --- */
function SignaturePad({ value, onChange, disabled }) {
  const ref = React.useRef(null);
  const drawing = React.useRef(false);
  const getPos = (e, canvas) => {
    const rect = canvas.getBoundingClientRect();
    const src = e.touches ? e.touches[0] : e;
    return [(src.clientX - rect.left) * (canvas.width / rect.width),
            (src.clientY - rect.top)  * (canvas.height / rect.height)];
  };
  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (value) {
      const img = new Image();
      img.onload = () => { ctx.clearRect(0, 0, canvas.width, canvas.height); ctx.drawImage(img, 0, 0); };
      img.src = value;
    } else {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }
  }, [value]);
  const start = (e) => {
    if (disabled) return;
    e.preventDefault();
    drawing.current = true;
    const canvas = ref.current;
    const ctx = canvas.getContext("2d");
    ctx.beginPath();
    ctx.moveTo(...getPos(e, canvas));
  };
  const move = (e) => {
    if (!drawing.current || disabled) return;
    e.preventDefault();
    const canvas = ref.current;
    const ctx = canvas.getContext("2d");
    ctx.lineWidth = 2; ctx.lineCap = "round"; ctx.strokeStyle = "#1a1a2e";
    ctx.lineTo(...getPos(e, canvas));
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(...getPos(e, canvas));
  };
  const end = () => {
    if (!drawing.current) return;
    drawing.current = false;
    if (onChange) onChange(ref.current.toDataURL());
  };
  const clear = () => {
    const canvas = ref.current;
    canvas.getContext("2d").clearRect(0, 0, canvas.width, canvas.height);
    if (onChange) onChange("");
  };
  return (
    <div>
      <canvas ref={ref} width={380} height={100}
        style={{ border: "1px solid var(--line)", borderRadius: 8, background: "#FAFAFE", cursor: disabled ? "default" : "crosshair", touchAction: "none", maxWidth: "100%" }}
        onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end}
        onTouchStart={start} onTouchMove={move} onTouchEnd={end}
      />
      {!disabled && <button className="btn btn-mini" style={{ marginTop: 4 }} onClick={clear}>Clear</button>}
    </div>
  );
}

/* --- Cold Room Temperature Check form --- */
const CR_ZONES = [
  { key: "receiving",     label: "Receiving Area",   range: "8–10°C" },
  { key: "blast",         label: "Blast Area",        range: "8–10°C" },
  { key: "packing",       label: "Packing Area",       range: "8–10°C" },
  { key: "dispatchBlast", label: "Dispatch Blast",     range: "0.5–2°C" },
  { key: "dispatchArea",  label: "Dispatch Area",      range: "0.5–2°C" },
];
const CR_TIMES = ["0800hrs", "1200hrs", "1600hrs"];

function newCrReading() {
  const r = { date: new Date().toISOString().slice(0, 10) };
  CR_ZONES.forEach((z) => CR_TIMES.forEach((t) => (r[z.key + "_" + t] = "")));
  r.doneBy = ""; r.verification = "";
  return r;
}
function newCrForm(me) {
  return {
    id: "cr-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
    type: "coldroom",
    week: isoWeek(new Date().toISOString().slice(0, 10)),
    readings: [newCrReading()],
    corrective: [],
    submitted: null, submittedBy: null,
  };
}

function ColdRoomForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  if (!form.readings || !Array.isArray(form.readings)) {
    return <p className="muted" style={{ padding: 16 }}>This record is missing readings data. Please delete it and submit a new check.</p>;
  }
  const setR = (i, k) => (e) => {
    const readings = form.readings.map((r2, j) => j === i ? { ...r2, [k]: e.target.value } : r2);
    onChange({ ...form, readings });
  };
  const addRow = () => onChange({ ...form, readings: [...form.readings, newCrReading()] });
  const setC = (i, k) => (e) => {
    const corrective = form.corrective.map((r2, j) => j === i ? { ...r2, [k]: e.target.value } : r2);
    onChange({ ...form, corrective });
  };
  const addCorrective = () => onChange({ ...form, corrective: [...form.corrective, { date: "", cr: "", oosTemp: "", action: "", retest: "", retestTime: "", doneBy: "", verification: "" }] });

  const inSpec = (zone, val) => {
    const v = parseFloat(val);
    if (isNaN(v) || val === "") return null;
    const hi = zone.key.startsWith("dispatch") ? 2 : 10;
    const lo = zone.key.startsWith("dispatch") ? 0.5 : 8;
    return v >= lo && v <= hi;
  };

  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Week</span>
          <input className="inp" type="number" value={form.week} onChange={(e) => onChange({ ...form, week: +e.target.value })} disabled={readOnly} /></label>
        <label className="fld"><span>IPF Ref</span><input className="inp" value="IPF 01" disabled /></label>
        <label className="fld"><span>Version</span><input className="inp" value="VERSION 2" disabled /></label>
      </div>

      <div className="tbl-wrap" style={{ marginBottom: 14 }}>
        <table style={{ width: "100%", fontSize: 12 }}>
          <thead>
            <tr>
              <th rowSpan={2} style={{ minWidth: 90 }}>Date</th>
              {CR_ZONES.map((z) => (
                <th key={z.key} colSpan={3} style={{ textAlign: "center" }}>
                  {z.label}<div className="muted small">{z.range}</div>
                </th>
              ))}
              <th rowSpan={2}>Done by</th>
              <th rowSpan={2}>Verification</th>
            </tr>
            <tr>
              {CR_ZONES.flatMap((z) => CR_TIMES.map((t) => <th key={z.key+t} className="r">{t}</th>))}
            </tr>
          </thead>
          <tbody>
            {form.readings.map((r2, i) => (
              <tr key={i}>
                <td>{readOnly ? dfmt(r2.date) : <input className="inp" type="date" value={r2.date} onChange={setR(i,"date")} style={{ minWidth: 120 }} />}</td>
                {CR_ZONES.flatMap((z) => CR_TIMES.map((t) => {
                  const k = z.key + "_" + t;
                  const spec = inSpec(z, r2[k]);
                  return (
                    <td key={k} className="r" style={spec === false ? { background: "#FDECEA" } : spec === true ? { background: "#EDF7EE" } : {}}>
                      {readOnly ? (r2[k] || "\u2014") : (
                        <input className="inp" type="number" step="0.1" value={r2[k]} onChange={setR(i,k)}
                          style={{ width: 62, padding: "3px 5px", fontSize: 11.5, textAlign: "right",
                            background: spec === false ? "#FDECEA" : spec === true ? "#EDF7EE" : "" }} />
                      )}
                    </td>
                  );
                }))}
                <td>{readOnly ? r2.doneBy : <input className="inp" value={r2.doneBy} onChange={setR(i,"doneBy")} style={{ minWidth: 90 }} />}</td>
                <td>{readOnly ? r2.verification : <input className="inp" value={r2.verification} onChange={setR(i,"verification")} style={{ minWidth: 90 }} />}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {!readOnly && <button className="btn btn-mini" style={{ marginBottom: 14 }} onClick={addRow}>+ Add day</button>}

      <h4 style={{ marginBottom: 8 }}>Corrective Action \u2014 Out of Spec Temperature</h4>
      <div className="tbl-wrap" style={{ marginBottom: 14 }}>
        <table style={{ width: "100%", fontSize: 12 }}>
          <thead><tr>
            <th>Date</th><th>CR #</th><th>OOS Temp °C</th><th>Corrective Action</th>
            <th>Temp Retest °C</th><th>Time of Retest</th><th>Done by</th><th>Verification</th>
          </tr></thead>
          <tbody>
            {form.corrective.map((c, i) => (
              <tr key={i}>
                {["date","cr","oosTemp","action","retest","retestTime","doneBy","verification"].map((k) => (
                  <td key={k}>{readOnly ? (c[k]||"\u2014") : <input className="inp" value={c[k]} onChange={setC(i,k)} style={{ minWidth: 70, fontSize: 11.5 }} />}</td>
                ))}
              </tr>
            ))}
            {!form.corrective.length && <tr><td colSpan={8} className="muted center">No out-of-spec entries.</td></tr>}
          </tbody>
        </table>
      </div>
      {!readOnly && <button className="btn btn-mini" style={{ marginBottom: 14 }} onClick={addCorrective}>+ Add corrective action</button>}

      <div style={{ marginBottom: 14 }}>
        <p className="muted small" style={{ marginBottom: 6 }}>Signature{readOnly ? "" : " — sign with your finger or mouse"}</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} />
      </div>
      {!readOnly && (
        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
          <button className="btn btn-primary" style={{ width: "auto" }} onClick={() => onSave(form)}>Submit &amp; save</button>
          <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
        </div>
      )}
    </div>
  );
}


/* ─────────────── CRATE CLEANING RECORD ─────────────── */
function newCrateForm(me) {
  return { id: "crate-" + Date.now(), type: "crate", month: new Date().toISOString().slice(0, 7),
    rows: [{ date: new Date().toISOString().slice(0, 10), buckets: "", crates: "", broken: "", cleanedBy: me ? me.name : "", verifiedBy: "" }],
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
function CrateForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  const setR = (i, k) => (e) => { const rows = form.rows.map((r, j) => j === i ? { ...r, [k]: e.target.value } : r); onChange({ ...form, rows }); };
  const addRow = () => onChange({ ...form, rows: [...form.rows, { date: new Date().toISOString().slice(0, 10), buckets: "", crates: "", broken: "", cleanedBy: "", verifiedBy: "" }] });
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Month</span><input className="inp" type="month" value={form.month} onChange={(e) => onChange({ ...form, month: e.target.value })} disabled={readOnly} /></label>
      </div>
      <div className="muted small" style={{ marginBottom: 10 }}>
        <b>Procedure:</b> Pre-dilute soap, soak crate/bucket in wash-tub for 5–8 min → Scrub → Rinse → Sanitise → Inspect and remove broken items.
      </div>
      <div className="tbl-wrap" style={{ marginBottom: 10 }}>
        <table><thead><tr>
          <th>Date</th><th className="r"># Reaping buckets cleaned</th><th className="r"># Crates/trays cleaned</th><th className="r"># Broken removed</th><th>Cleaned by</th><th>Verified by</th>
        </tr></thead><tbody>
          {form.rows.map((r, i) => (
            <tr key={i}>
              <td>{readOnly ? dfmt(r.date) : <input className="inp" type="date" value={r.date} onChange={setR(i,"date")} style={{ minWidth: 120 }} />}</td>
              {["buckets","crates","broken"].map(k => <td key={k} className="r">{readOnly ? r[k] || "—" : <input className="inp" type="number" min="0" value={r[k]} onChange={setR(i,k)} style={{ width: 70, textAlign: "right" }} />}</td>)}
              {["cleanedBy","verifiedBy"].map(k => <td key={k}>{readOnly ? r[k] || "—" : <input className="inp" value={r[k]} onChange={setR(i,k)} style={{ minWidth: 90 }} />}</td>)}
            </tr>
          ))}
        </tbody></table>
      </div>
      {!readOnly && <button className="btn btn-mini" style={{ marginBottom: 14 }} onClick={addRow}>+ Add row</button>}
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8 }}><button className="btn btn-primary" style={{ width:"auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
  );
}

/* ─────────────── LABELLING ROLL APPROVAL ─────────────── */
function newLabelForm(me) {
  return { id: "label-" + Date.now(), type: "label", date: new Date().toISOString().slice(0, 10),
    customer: "", packType: "Punnets",
    punnetFirst: "", punnetLast: "", boxFirst: "", boxLast: "",
    doneBy: me ? me.name : "", supervisorSign: "",
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
/* ─────────────── PHOTO CAPTURE FIELD ─────────────── */
function PhotoCapture({ value, onChange, disabled, label }) {
  const inputRef = React.useRef(null);
  const toDataUrl = (file) => new Promise((res) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.readAsDataURL(file);
  });
  const onFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    // resize to max 1200px wide to keep storage lean
    const img = new Image();
    img.src = URL.createObjectURL(file);
    img.onload = () => {
      const MAX = 1200;
      const scale = img.width > MAX ? MAX / img.width : 1;
      const canvas = document.createElement("canvas");
      canvas.width = img.width * scale; canvas.height = img.height * scale;
      canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
      onChange(canvas.toDataURL("image/jpeg", 0.82));
      URL.revokeObjectURL(img.src);
    };
    e.target.value = "";
  };
  return (
    <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
      {value
        ? <img src={value} alt={label} style={{ maxWidth:160, maxHeight:100, borderRadius:6, border:"1px solid var(--line)", objectFit:"contain", background:"#FAFAFE" }} />
        : <div style={{ width:160, height:80, borderRadius:6, border:"1px dashed var(--line)", display:"flex", alignItems:"center", justifyContent:"center", color:"var(--muted)", fontSize:11 }}>No photo</div>}
      {!disabled && (
        <div style={{ display:"flex", gap:6 }}>
          <button type="button" className="btn btn-mini" onClick={() => inputRef.current && inputRef.current.click()}>
            {value ? "\ud83d\udcf7 Retake" : "\ud83d\udcf7 Take photo"}
          </button>
          {value && <button type="button" className="btn btn-mini" onClick={() => onChange("")}>Clear</button>}
          <input ref={inputRef} type="file" accept="image/*" capture="environment" style={{ display:"none" }} onChange={onFile} />
        </div>
      )}
    </div>
  );
}

function LabelForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  const setPhoto = (k) => (v) => onChange({ ...form, [k]: v });
  const PACK_TYPES = ["Punnets", "Tubes", "Bulk", "Buckets"];
  const PHOTO_ROWS = [
    [form.packType || "Pack Type", "punnetFirst", "punnetLast", "punnetFirstPhoto", "punnetLastPhoto"],
    ["Box end label",              "boxFirst",    "boxLast",    "boxFirstPhoto",    "boxLastPhoto"],
  ];
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Date</span><input className="inp" type="date" value={form.date} onChange={set("date")} disabled={readOnly} /></label>
        <label className="fld"><span>Customer</span><input className="inp" value={form.customer || ""} onChange={set("customer")} disabled={readOnly} placeholder="Customer name" /></label>
        <label className="fld"><span>Pack type</span>
          {readOnly
            ? <span className="inp" style={{ display:"block" }}>{form.packType || "—"}</span>
            : <select className="inp" value={form.packType || "Punnets"} onChange={set("packType")}>
                {PACK_TYPES.map((t) => <option key={t}>{t}</option>)}
              </select>}
        </label>
        <label className="fld"><span>Done by</span><input className="inp" value={form.doneBy} onChange={set("doneBy")} disabled={readOnly} /></label>
        <label className="fld"><span>Supervisor sign-off</span><input className="inp" value={form.supervisorSign} onChange={set("supervisorSign")} disabled={readOnly} /></label>
      </div>
      <div style={{ marginBottom: 14 }}>
        {PHOTO_ROWS.map(([rowLabel, fk, lk, fph, lph]) => (
          <div key={rowLabel} style={{ marginBottom: 16, padding: 12, background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 10 }}>
            <div className="strong" style={{ marginBottom: 10, fontSize: 13 }}>{rowLabel}</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <div>
                <div className="muted small" style={{ marginBottom: 6 }}>First roll label</div>
                <PhotoCapture value={form[fph] || ""} onChange={setPhoto(fph)} disabled={readOnly} label={"First " + rowLabel} />
                <input className="inp" style={{ marginTop: 6, fontSize: 12 }} placeholder="Label reference / notes" value={form[fk] || ""} onChange={set(fk)} disabled={readOnly} />
              </div>
              <div>
                <div className="muted small" style={{ marginBottom: 6 }}>Last roll label</div>
                <PhotoCapture value={form[lph] || ""} onChange={setPhoto(lph)} disabled={readOnly} label={"Last " + rowLabel} />
                <input className="inp" style={{ marginTop: 6, fontSize: 12 }} placeholder="Label reference / notes" value={form[lk] || ""} onChange={set(lk)} disabled={readOnly} />
              </div>
            </div>
          </div>
        ))}
      </div>
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8 }}><button className="btn btn-primary" style={{ width:"auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
  );
}

/* ─────────────── PALLET COMPOSITION ─────────────── */
function newPalletForm(me) {
  return { id: "pallet-" + Date.now(), type: "pallet", date: new Date().toISOString().slice(0, 10),
    palletRef: "", customer: "", totalBoxes: "", dispatchTemp: "", qcCheck: "", doneBy: me ? me.name : "",
    lines: [{ variety: "", block: "", runNumber: "", boxes: "" }],
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
function PalletForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  const setL = (i, k) => (e) => { const lines = form.lines.map((l, j) => j === i ? { ...l, [k]: e.target.value } : l); onChange({ ...form, lines }); };
  const addLine = () => onChange({ ...form, lines: [...form.lines, { variety: "", block: "", runNumber: "", boxes: "" }] });
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Date</span><input className="inp" type="date" value={form.date} onChange={set("date")} disabled={readOnly} /></label>
        <label className="fld"><span>Pallet reference</span><input className="inp" value={form.palletRef} onChange={set("palletRef")} disabled={readOnly} /></label>
        <label className="fld"><span>Customer</span><input className="inp" value={form.customer} onChange={set("customer")} disabled={readOnly} /></label>
        <label className="fld"><span>Total number of boxes</span><input className="inp" type="number" value={form.totalBoxes} onChange={set("totalBoxes")} disabled={readOnly} /></label>
        <label className="fld"><span>Dispatch temperature</span><input className="inp" value={form.dispatchTemp} onChange={set("dispatchTemp")} disabled={readOnly} /></label>
        <label className="fld"><span>QC check</span><input className="inp" value={form.qcCheck} onChange={set("qcCheck")} disabled={readOnly} /></label>
        <label className="fld"><span>Done by</span><input className="inp" value={form.doneBy} onChange={set("doneBy")} disabled={readOnly} /></label>
      </div>
      <div className="tbl-wrap" style={{ marginBottom: 10 }}>
        <table><thead><tr><th>Variety</th><th>Block</th><th>Run number</th><th className="r">Number of boxes</th></tr></thead>
        <tbody>
          {form.lines.map((l, i) => (
            <tr key={i}>
              {["variety","block","runNumber"].map(k => <td key={k}>{readOnly ? l[k] || "—" : <input className="inp" value={l[k]} onChange={setL(i,k)} />}</td>)}
              <td className="r">{readOnly ? l.boxes || "—" : <input className="inp" type="number" value={l.boxes} onChange={setL(i,"boxes")} style={{ width: 80, textAlign: "right" }} />}</td>
            </tr>
          ))}
        </tbody></table>
      </div>
      {!readOnly && <button className="btn btn-mini" style={{ marginBottom: 14 }} onClick={addLine}>+ Add line</button>}
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8 }}><button className="btn btn-primary" style={{ width:"auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
  );
}

/* ─────────────── PERSONAL HYGIENE (shared) ─────────────── */
const HYGIENE_FIELD_ITEMS = [
  { section: "Personal Hygiene", items: ["Finger nails clean and neat","No jewellery","Cuts and wounds protected","All employees healthy","No allergens"] },
  { section: "Protective Clothing", items: ["Uniform correct for function","Uniform clean","Hairnet worn","No undergarments showing","Visitors correctly dressed"] },
  { section: "Handwashing and Sanitising", items: ["Staff wash hands as required","Correct / frequent use of sanitiser"] },
  { section: "Work Areas and General Fruit Handling", items: ["No personal belongings","No eating / drinking / smoking","No spitting at workplace"] },
  { section: "Toilets", items: ["Toilets cleaned and tidy","Soap and sanitiser present"] },
  { section: "Trailer and Equipment Check", items: ["Trailer cleaned and tidy","Sanitiser and soap present","Punnets and buckets cleaned"] },
  { section: "Vehicle Check", items: ["Vehicle clean and tidy","Loaders wearing PPE","No. of people per reaping station"] },
];
const HYGIENE_PACK_ITEMS = [
  { section: "Personal Hygiene", items: ["Finger nails clean and neat","No jewellery","Cuts and wounds protected","Gloves applied","Employees health","Allergens"] },
  { section: "Protective Clothing", items: ["Uniform correct for function","Uniform clean","Hairnet worn","No undergarments showing","Visitors correctly dressed"] },
  { section: "Handwashing and Sanitising", items: ["Staff wash hands as required","Correct / frequent use of sanitiser"] },
  { section: "Work Areas and General Fruit Handling", items: ["No personal belongings","No eating / drinking / gum","Untidy / unattended work area","No spitting at workplace"] },
];
function makeHygieneForm(type, me) {
  const sections = type === "hygieneField" ? HYGIENE_FIELD_ITEMS : HYGIENE_PACK_ITEMS;
  const items = sections.flatMap(s => s.items.map(item => ({ section: s.section, item, status: "", comments: "" })));
  return { id: type + "-" + Date.now(), type, date: new Date().toISOString().slice(0, 10),
    time: "", doneBy: me ? me.name : "", verifiedBy: "", overallComment: "", items,
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
function HygieneForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  if (!form.items || !Array.isArray(form.items)) {
    return <p className="muted" style={{ padding: 16 }}>This record is missing checklist data. Please delete it and submit a new inspection.</p>;
  }
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  const setItem = (i, k) => (e) => { const items = form.items.map((x, j) => j === i ? { ...x, [k]: e.target.value } : x); onChange({ ...form, items }); };
  const sections = [...new Set(form.items.map(x => x.section))];
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Date</span><input className="inp" type="date" value={form.date} onChange={set("date")} disabled={readOnly} /></label>
        <label className="fld"><span>Time</span><input className="inp" type="time" value={form.time} onChange={set("time")} disabled={readOnly} /></label>
        <label className="fld"><span>Done by</span><input className="inp" value={form.doneBy} onChange={set("doneBy")} disabled={readOnly} /></label>
        <label className="fld"><span>Verified by</span><input className="inp" value={form.verifiedBy} onChange={set("verifiedBy")} disabled={readOnly} /></label>
      </div>
      {sections.map(sec => (
        <div key={sec} style={{ marginBottom: 12 }}>
          <div className="dayhead" style={{ borderRadius: 6, marginBottom: 0 }}>
            <table style={{ width:"100%", borderCollapse:"collapse" }}><tbody>
              <tr><td style={{ padding:"5px 10px", color:"#fff", fontWeight:700, fontSize:11, textTransform:"uppercase", letterSpacing:".06em" }}>{sec}</td></tr>
            </tbody></table>
          </div>
          <div className="tbl-wrap"><table style={{ width:"100%" }}>
            <thead><tr><th>Item</th><th style={{ width:140 }}>Status</th><th>Comments</th></tr></thead>
            <tbody>
              {form.items.filter(x => x.section === sec).map((x, _) => {
                const i = form.items.indexOf(x);
                return (
                  <tr key={x.item} className={!x.status && !readOnly ? "hidef" : ""}>
                    <td>{x.item}</td>
                    <td>{readOnly ? <span className={"badge " + (x.status === "Pass" ? "ok" : x.status === "Fail" ? "err" : "")}>{x.status || "—"}</span> :
                      <select className="inp" style={{ padding:"3px 6px", fontSize:12 }} value={x.status} onChange={setItem(i,"status")}>
                        <option value="">Select…</option><option>Pass</option><option>Fail</option><option>N/A</option>
                      </select>}</td>
                    <td>{readOnly ? x.comments || "—" : <input className="inp" style={{ fontSize:12 }} value={x.comments} onChange={setItem(i,"comments")} placeholder="Optional" />}</td>
                  </tr>
                );
              })}
            </tbody>
          </table></div>
        </div>
      ))}
      <label className="fld" style={{ marginBottom: 14 }}><span>Overall comment</span>
        <input className="inp" value={form.overallComment || ""} onChange={set("overallComment")} disabled={readOnly} /></label>
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8, marginTop:8 }}>
        <button className="btn btn-primary" style={{ width:"auto" }} disabled={form.items.some(x => !x.status)} onClick={() => onSave(form)}>Submit &amp; save</button>
        <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
        {form.items.some(x => !x.status) && <span className="muted small" style={{ alignSelf:"center" }}>Complete all status fields to submit.</span>}
      </div>}
    </div>
  );
}

/* ─────────────── POST-MACHINE QC REPORT ─────────────── */
const QC_FIELDS = [
  "Underripe / Red bum","Overripe / Soft","Collapsed and bleeding","Minor shrivel","Major shrivel",
  "Pick scars, punctures and tears","Blemishes","Stem retention / Flowers and beards","Rot / Mould",
  "Pest presence","Insect damage","Fruit size out of spec",
];
const QC_SCORED = [
  { key: "brix", label: "Brix reading" },
  { key: "flavour", label: "Flavour / Aroma (1=poor, 2=acceptable, 3=good)" },
  { key: "bloom", label: "Bloom (1=poor, 2=acceptable, 3=good)" },
  { key: "foreignMatter", label: "Foreign matter (1=high, 2=low, 3=none)" },
  { key: "pulpTemp", label: "Pulp temperature (°C)" },
  { key: "agrostaAvg", label: "Agrosta Average" },
  { key: "agrostaStd", label: "Agrosta STD" },
  { key: "rejectsAvg", label: "Rejects Average" },
  { key: "rejectsStd", label: "Rejects STD" },
];
function newQcForm(me) {
  const defects = {}; QC_FIELDS.forEach(f => (defects[f] = ""));
  return { id: "qc-" + Date.now(), type: "qcReport", date: new Date().toISOString().slice(0, 10),
    sampleSize: "100", ...defects, ...Object.fromEntries(QC_SCORED.map(x => [x.key, ""])),
    qualityController: "", seniorQC: "", machineOperator: "", qualityCoordinator: me ? me.name : "",
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
function QcForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Date</span><input className="inp" type="date" value={form.date} onChange={set("date")} disabled={readOnly} /></label>
        <label className="fld"><span>Sample size</span><input className="inp" value={form.sampleSize} onChange={set("sampleSize")} disabled={readOnly} /></label>
      </div>
      <h4 style={{ marginBottom: 8 }}>Defects (count out of sample)</h4>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        {QC_FIELDS.map(f => (
          <label key={f} className="fld"><span style={{ fontSize: 11 }}>{f}</span>
            <input className="inp" type="number" min="0" value={form[f] || ""} onChange={set(f)} disabled={readOnly} /></label>
        ))}
      </div>
      <h4 style={{ marginBottom: 8 }}>Quality scores &amp; readings</h4>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        {QC_SCORED.map(x => (
          <label key={x.key} className="fld"><span style={{ fontSize: 11 }}>{x.label}</span>
            <input className="inp" type="number" step="0.1" value={form[x.key] || ""} onChange={set(x.key)} disabled={readOnly} /></label>
        ))}
      </div>
      <h4 style={{ marginBottom: 8 }}>Sign-offs</h4>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        {[["qualityController","Quality Controller"],["seniorQC","Senior Quality Controller"],["machineOperator","Machine Operator"],["qualityCoordinator","Quality Coordinator"]].map(([k,l]) => (
          <label key={k} className="fld"><span>{l}</span><input className="inp" value={form[k] || ""} onChange={set(k)} disabled={readOnly} /></label>
        ))}
      </div>
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8 }}><button className="btn btn-primary" style={{ width:"auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
  );
}

/* ─────────────── SCALE CALIBRATION ─────────────── */
const SCALE_WEIGHTS = ["100g","200g","500g","2kg","5kg"];
const SCALE_TIMES = ["8:00 AM","12:00 PM"];
function newScaleForm(me) {
  const readings = {};
  SCALE_TIMES.forEach(t => SCALE_WEIGHTS.forEach(w => (readings[t + "_" + w] = "")));
  return { id: "scale-" + Date.now(), type: "scale", date: new Date().toISOString().slice(0, 10),
    scaleNo: "", ...readings, doneBy: me ? me.name : "", verifiedBy: "",
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
function ScaleForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Date</span><input className="inp" type="date" value={form.date} onChange={set("date")} disabled={readOnly} /></label>
        <label className="fld"><span>Scale No.</span><input className="inp" value={form.scaleNo} onChange={set("scaleNo")} disabled={readOnly} /></label>
        <label className="fld"><span>Done by</span><input className="inp" value={form.doneBy} onChange={set("doneBy")} disabled={readOnly} /></label>
        <label className="fld"><span>Verified by</span><input className="inp" value={form.verifiedBy} onChange={set("verifiedBy")} disabled={readOnly} /></label>
      </div>
      <div className="tbl-wrap" style={{ marginBottom: 14 }}>
        <table><thead>
          <tr><th rowSpan={2}>Time</th>{SCALE_WEIGHTS.map(w => <th key={w} className="r">{w}</th>)}</tr>
        </thead><tbody>
          {SCALE_TIMES.map(t => (
            <tr key={t}><td className="strong">{t}</td>
              {SCALE_WEIGHTS.map(w => {
                const k = t + "_" + w; const v = form[k] || "";
                const exp = parseFloat(w); const got = parseFloat(v);
                const ok = !v ? null : Math.abs(got - exp) / exp < 0.02; // 2% tolerance
                return <td key={w} className="r" style={ok === false ? { background:"#FDECEA" } : ok === true ? { background:"#EDF7EE" } : {}}>
                  {readOnly ? v || "—" : <input className="inp" type="number" step="0.1" value={v} onChange={set(k)} style={{ width:70, textAlign:"right", background: ok === false ? "#FDECEA" : ok === true ? "#EDF7EE" : "" }} />}
                </td>;
              })}
            </tr>
          ))}
        </tbody></table>
      </div>
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8 }}><button className="btn btn-primary" style={{ width:"auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
  );
}

/* ─────────────── TRUCK HYGIENE & TEMP CHECK ─────────────── */
function newTruckForm(me) {
  return { id: "truck-" + Date.now(), type: "truck",
    month: new Date().toISOString().slice(0, 7),
    rows: [{ date: new Date().toISOString().slice(0, 10), truckReg: "", driverSign: "", tempSetting: "", loadingTemp: "", hygiene: "Satisfactory", comment: "", corrective: "" }],
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
function TruckForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  const setR = (i, k) => (e) => { const rows = form.rows.map((r, j) => j === i ? { ...r, [k]: e.target.value } : r); onChange({ ...form, rows }); };
  const addRow = () => onChange({ ...form, rows: [...form.rows, { date: new Date().toISOString().slice(0, 10), truckReg: "", driverSign: "", tempSetting: "", loadingTemp: "", hygiene: "Satisfactory", comment: "", corrective: "" }] });
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Month</span><input className="inp" type="month" value={form.month} onChange={(e) => onChange({ ...form, month: e.target.value })} disabled={readOnly} /></label>
      </div>
      <div className="muted small" style={{ marginBottom: 8 }}>Hygiene status: Satisfactory · Unsatisfactory · Not Cleaned. Temp setting: +2°C continuously. Loading temp: 1–4°C.</div>
      <div className="tbl-wrap" style={{ marginBottom: 10 }}>
        <table><thead><tr>
          <th>Date</th><th>Truck Reg</th><th>Driver sign</th><th>Temp setting (°C)</th><th>Loading temp (°C)</th><th>Hygiene (S/U/A)</th><th>Comment</th><th>Corrective action</th>
        </tr></thead><tbody>
          {form.rows.map((r, i) => (
            <tr key={i} className={(r.hygiene === "U" || r.hygiene === "Unsatisfactory" || r.hygiene === "A" || r.hygiene === "Not Cleaned") ? "hidef" : ""}>
              <td>{readOnly ? dfmt(r.date) : <input className="inp" type="date" value={r.date} onChange={setR(i,"date")} style={{ minWidth:120 }} />}</td>
              {["truckReg","driverSign"].map(k => <td key={k}>{readOnly ? r[k] || "—" : <input className="inp" value={r[k]} onChange={setR(i,k)} style={{ minWidth:80 }} />}</td>)}
              {["tempSetting","loadingTemp"].map(k => <td key={k} className="r">{readOnly ? r[k] || "—" : <input className="inp" type="number" step="0.1" value={r[k]} onChange={setR(i,k)} style={{ width:70, textAlign:"right" }} />}</td>)}
              <td>{readOnly ? r.hygiene || "—" : <select className="inp" style={{ padding:"3px 6px" }} value={r.hygiene} onChange={setR(i,"hygiene")}><option value="S">Satisfactory</option><option value="U">Unsatisfactory</option><option value="A">Not Cleaned</option></select>}</td>
              {["comment","corrective"].map(k => <td key={k}>{readOnly ? r[k] || "—" : <input className="inp" value={r[k]} onChange={setR(i,k)} style={{ minWidth:90 }} />}</td>)}
            </tr>
          ))}
        </tbody></table>
      </div>
      {!readOnly && <button className="btn btn-mini" style={{ marginBottom: 14 }} onClick={addRow}>+ Add row</button>}
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8 }}><button className="btn btn-primary" style={{ width:"auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
  );
}

/* ─────────────── VISITOR QUESTIONNAIRE ─────────────── */
function newVisitorForm(me) {
  return { id: "visitor-" + Date.now(), type: "visitor", date: new Date().toISOString().slice(0, 10),
    visitorName: "", company: "", purpose: "", hostName: me ? me.name : "",
    sickLast48h: "", hasWounds: "", wearingJewellery: "", hasAllergies: "", allergyDetails: "",
    agreedToPolicy: false, timeIn: "", timeOut: "",
    signature: "", submitted: null, submittedBy: null, submittedUsername: null };
}
function VisitorForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  const setB = (k) => (e) => onChange({ ...form, [k]: e.target.checked });
  const yn = (k, label) => (
    <label key={k} className="fld"><span>{label}</span>
      {readOnly ? <span className={"badge " + (form[k] === "Yes" ? "err" : "ok")}>{form[k] || "—"}</span> :
        <select className="inp" value={form[k] || ""} onChange={set(k)}>
          <option value="">Select…</option><option>Yes</option><option>No</option>
        </select>}
    </label>
  );
  return (
    <div>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Date</span><input className="inp" type="date" value={form.date} onChange={set("date")} disabled={readOnly} /></label>
        <label className="fld"><span>Visitor name</span><input className="inp" value={form.visitorName} onChange={set("visitorName")} disabled={readOnly} /></label>
        <label className="fld"><span>Company</span><input className="inp" value={form.company} onChange={set("company")} disabled={readOnly} /></label>
        <label className="fld"><span>Purpose of visit</span><input className="inp" value={form.purpose} onChange={set("purpose")} disabled={readOnly} /></label>
        <label className="fld"><span>Host name</span><input className="inp" value={form.hostName} onChange={set("hostName")} disabled={readOnly} /></label>
        <label className="fld"><span>Time in</span><input className="inp" type="time" value={form.timeIn} onChange={set("timeIn")} disabled={readOnly} /></label>
        <label className="fld"><span>Time out</span><input className="inp" type="time" value={form.timeOut} onChange={set("timeOut")} disabled={readOnly} /></label>
      </div>
      <h4 style={{ marginBottom: 10 }}>Health &amp; safety declaration</h4>
      <div className="fgrid" style={{ marginBottom: 14 }}>
        {yn("sickLast48h", "Sick in the last 48 hours?")}
        {yn("hasWounds", "Open cuts or wounds?")}
        {yn("wearingJewellery", "Wearing jewellery?")}
        {yn("hasAllergies", "Known food allergies?")}
        {form.hasAllergies === "Yes" && <label className="fld"><span>Allergy details</span><input className="inp" value={form.allergyDetails || ""} onChange={set("allergyDetails")} disabled={readOnly} /></label>}
      </div>
      <label style={{ display:"flex", gap:8, alignItems:"center", marginBottom:14, fontSize:13 }}>
        <input type="checkbox" checked={!!form.agreedToPolicy} onChange={setB("agreedToPolicy")} disabled={readOnly} />
        Visitor has read and agreed to the site food safety &amp; hygiene policy
      </label>
      <div style={{ marginBottom: 14 }}><p className="muted small" style={{ marginBottom: 6 }}>Visitor signature</p>
        <SignaturePad value={form.signature || ""} onChange={readOnly ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} /></div>
      {!readOnly && <div style={{ display:"flex", gap:8 }}><button className="btn btn-primary" style={{ width:"auto" }} disabled={!form.agreedToPolicy} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
  );
}

/* ================= COMPLIANCE ================= */
const COMPLIANCE_KEY = "ml-compliance-v1";
const ELIFAB_ITEMS = [
    { code: "1.1", section: "1. Accumulation Module", desc: "Clean and intact feed-on chute", status: "", comments: "" },
    { code: "1.2", section: "1. Accumulation Module", desc: "Clean and intact belt", status: "", comments: "" },
    { code: "1.3", section: "1. Accumulation Module", desc: "Clean and intact walls of accumulation module", status: "", comments: "" },
    { code: "1.4", section: "1. Accumulation Module", desc: "Clean and intact feed-on curtain and frame", status: "", comments: "" },
    { code: "2.1", section: "2. Singulation/Acceleration Modules", desc: "Remove and clean lane divider covers", status: "", comments: "" },
    { code: "2.2", section: "2. Singulation/Acceleration Modules", desc: "Clean belts on singulation and acceleration modules", status: "", comments: "" },
    { code: "2.3", section: "2. Singulation/Acceleration Modules", desc: "Clean pull-out tray and interior chutes", status: "", comments: "" },
    { code: "2.4", section: "2. Singulation/Acceleration Modules", desc: "Remove any built-up debris or berries from scraper plate", status: "", comments: "" },
    { code: "3.1", section: "3. Carrier Module", desc: "Clean dark plates", status: "", comments: "" },
    { code: "3.2", section: "3. Carrier Module", desc: "Check timing belts and clean if required", status: "", comments: "" },
    { code: "3.3", section: "3. Carrier Module", desc: "Clean under the fruit sensors to ensure no build-up of berries", status: "", comments: "" },
    { code: "3.4", section: "3. Carrier Module", desc: "Clean feed-on chute", status: "", comments: "" },
    { code: "3.5", section: "3. Carrier Module", desc: "Clean feed-off chute", status: "", comments: "" },
    { code: "3.6", section: "3. Carrier Module", desc: "Clean lift-off lane divider covers", status: "", comments: "" },
    { code: "3.7", section: "3. Carrier Module", desc: "Clean rollers using the CIP bath", status: "", comments: "" },
    { code: "3.8", section: "3. Carrier Module", desc: "Hose inside and outside of catch pans", status: "", comments: "" },
    { code: "4.1", section: "4. Tech Module", desc: "Check camera windows are clean &amp; intact", status: "", comments: "" },
    { code: "4.2", section: "4. Tech Module", desc: "Ensure light ring plastic is clean", status: "", comments: "" },
    { code: "4.3", section: "4. Tech Module", desc: "All control buttons of the machine intact", status: "", comments: "" },
    { code: "5.1", section: "5. Manifolds", desc: "Clean sides of manifolds to remove any build-up of debris/berries", status: "", comments: "" },
    { code: "5.2", section: "5. Manifolds", desc: "Clean ejection curtains", status: "", comments: "" },
    { code: "5.3", section: "5. Manifolds", desc: "Clean jet outlets on manifold", status: "", comments: "" },
    { code: "6.1", section: "6. Selection Conveyors", desc: "Clean feed-off chute", status: "", comments: "" },
    { code: "6.2", section: "6. Selection Conveyors", desc: "Clean belt", status: "", comments: "" },
    { code: "6.3", section: "6. Selection Conveyors", desc: "Clean angled surfaces on and between conveyors", status: "", comments: "" },
    { code: "6.4", section: "6. Selection Conveyors", desc: "Clean backstop curtain", status: "", comments: "" },
    { code: "7.1", section: "7. Recirculation System", desc: "Clean belts on recirculation conveyors", status: "", comments: "" },
    { code: "7.2", section: "7. Recirculation System", desc: "Clean feed-off chute", status: "", comments: "" },
    { code: "7.3", section: "7. Recirculation System", desc: "Clean recirculation lift belt flights and belt", status: "", comments: "" },
    { code: "7.4", section: "7. Recirculation System", desc: "Clean recirculation lift belt curtain and split belt", status: "", comments: "" },
    { code: "8.1", section: "8. Final Checks", desc: "Check all the pull-out trays are clean", status: "", comments: "" },
    { code: "8.2", section: "8. Final Checks", desc: "Ensure the chain lubricators are not empty", status: "", comments: "" },
    { code: "8.3", section: "8. Final Checks", desc: "Ensure the chain has been run for a further 20 minutes after cleaning to allow for drying and relubrication", status: "", comments: "" },
  ];

function newElifabForm(me) {
  return {
    id: "cf-" + Date.now() + "-" + Math.floor(Math.random() * 1e6),
    type: "elifab",
    date: new Date().toISOString().slice(0, 10),
    week: isoWeek(new Date().toISOString().slice(0, 10)),
    completedBy: me ? me.name : "",
    verifiedBy: "",
    items: ELIFAB_ITEMS.map((x) => ({ ...x })),
    submitted: null,
    submittedBy: null,
  };
}

function ElifabForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const disabled = readOnly;
  // guard: older saved records may be missing items array
  if (!form.items || !Array.isArray(form.items)) {
    return <p className="muted" style={{ padding: 16 }}>This record is missing checklist data and cannot be displayed. Please delete it and submit a new inspection.</p>;
  }
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  const setItem = (i, k) => (e) => {
    const items = form.items.map((x, j) => j === i ? { ...x, [k]: e.target.value } : x);
    onChange({ ...form, items });
  };
  const sections = [...new Set(form.items.map((x) => x.section))];
  const allDone = Array.isArray(form.items) && form.items.every((x) => x.status);

  return (
    <div className="compliance-form">
      <div className="fgrid" style={{ marginBottom: 14 }}>
        <label className="fld"><span>Date</span>
          <input className="inp" type="date" value={form.date} onChange={set("date")} disabled={readOnly} /></label>
        <label className="fld"><span>Completed by</span>
          <input className="inp" value={form.completedBy} onChange={set("completedBy")} disabled={readOnly} /></label>
        <label className="fld"><span>Verified by</span>
          <input className="inp" value={form.verifiedBy} onChange={set("verifiedBy")} disabled={readOnly} /></label>
      </div>
      <div style={{ marginBottom: 14 }}>
        <p className="muted small" style={{ marginBottom: 6 }}>Signature{disabled ? "" : " — sign with your finger or mouse"}</p>
        <SignaturePad value={form.signature || ""} onChange={disabled ? null : (v) => onChange({ ...form, signature: v })} disabled={readOnly} />
      </div>
      {sections.map((sec) => (
        <div key={sec} style={{ marginBottom: 14 }}>
          <div className="dayhead" style={{ borderRadius: 6, marginBottom: 0 }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}><tbody>
              <tr><td style={{ padding: "5px 10px", color: "#fff", fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }}>{sec}</td></tr>
            </tbody></table>
          </div>
          <div className="tbl-wrap">
            <table style={{ width: "100%" }}>
              <thead><tr>
                <th style={{ width: 40 }}>#</th>
                <th>Checklist item</th>
                <th style={{ width: 160 }}>Status</th>
                <th>Comments / Corrective action</th>
              </tr></thead>
              <tbody>
                {form.items.filter((x) => x.section === sec).map((x, _) => {
                  const i = form.items.indexOf(x);
                  return (
                    <tr key={x.code} className={!x.status && !readOnly ? "hidef" : ""}>
                      <td className="mono muted">{x.code}</td>
                      <td>{x.desc.replace("&amp;", "&")}</td>
                      <td>
                        {readOnly ? <span className={"badge " + (x.status === "Pass" ? "ok" : x.status === "Fail" ? "err" : "")}>{x.status || "\u2014"}</span> : (
                          <select className="inp" style={{ padding: "3px 6px", fontSize: 12 }} value={x.status} onChange={setItem(i, "status")}>
                            <option value="">Select\u2026</option>
                            <option>Pass</option>
                            <option>Fail</option>
                            <option>N/A</option>
                          </select>
                        )}
                      </td>
                      <td>
                        {readOnly ? x.comments || "\u2014" : (
                          <input className="inp" style={{ fontSize: 12 }} value={x.comments} onChange={setItem(i, "comments")} placeholder="Optional" />
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      ))}
      {!readOnly && (
        <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
          <button className="btn btn-primary" style={{ width: "auto" }} disabled={!allDone} onClick={() => onSave(form)}>
            Submit &amp; save
          </button>
          <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          {!allDone && <span className="muted small" style={{ alignSelf: "center" }}>Complete all status fields to submit.</span>}
        </div>
      )}
    </div>
  );
}

/* ================= MASTER FILES ================= */
const MF_KEYS = {
  sections: "ml-master-sections-v1",
  commodities: "ml-master-commodities-v1",
  varieties: "ml-master-varieties-v1",
  blocks: "ml-master-blocks-v1",
  pallets: "ml-master-pallets-v1",
  crates: "ml-master-crates-v1",
  grades: "ml-master-grades-v1",
  seasons: "ml-master-seasons-v1",
  channels: "ml-master-channels-v1",
  transport: "ml-master-transport-v1",
  packtypes: "ml-master-packtypes-v1",
  sizes: "ml-master-sizes-v1",
  brands: "ml-master-brands-v1",
  invcodes: "ml-master-invcodes-v1",
  customers: "ml-master-customers-v1",
};

function MasterFiles({ me }) {
  // Admin only access
  const isAdmin = me?.role === "admin";
  
  const [tab, setTab] = useState("sections");
  const [data, setData] = useState({
    sections: [],
    commodities: [],
    varieties: [],
    blocks: [],
    pallets: [],
    crates: [],
    grades: [],
    seasons: [],
    channels: [],
    transport: [],
    packtypes: [],
    sizes: [],
    brands: [],
    invcodes: [],
    customers: [],
  });
  const [loading, setLoading] = useState(true);
  const [editForm, setEditForm] = useState(null);
  const [msg, setMsg] = useState(null);

  useEffect(() => {
    const loadAll = async () => {
      try {
        const loaded = {};
        for (const [key, storageKey] of Object.entries(MF_KEYS)) {
          const r = await loadShared(storageKey);
          loaded[key] = r ? JSON.parse(r) : [];
        }
        setData(loaded);
      } catch (e) {
        console.error("Load error:", e);
      }
      setLoading(false);
    };
    loadAll();
  }, []);

  const saveTab = async (key, items) => {
    const ok = await saveShared(MF_KEYS[key], JSON.stringify(items));
    return ok;
  };

  // COMMODITIES
  const addCommodity = () => {
    setEditForm({ type: "commodity", data: { id: "com-" + Date.now(), name: "", description: "" } });
  };
  const saveCommodity = async (commodity) => {
    const next = data.commodities.some((c) => c.id === commodity.id)
      ? data.commodities.map((c) => c.id === commodity.id ? commodity : c)
      : [...data.commodities, commodity];
    const ok = await saveTab("commodities", next);
    setData({ ...data, commodities: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Commodity saved."] : ["warn", "Saved locally."]);
  };
  const deleteCommodity = async (id) => {
    if (!window.confirm("Delete this commodity?")) return;
    const next = data.commodities.filter((c) => c.id !== id);
    const ok = await saveTab("commodities", next);
    setData({ ...data, commodities: next });
    setMsg(ok ? ["ok", "Commodity deleted."] : ["warn", "Deleted locally."]);
  };

  // SECTIONS
  const addSection = () => {
    setEditForm({ type: "section", data: { id: "sec-" + Date.now(), name: "", description: "", area: "" } });
  };
  const saveSection = async (section) => {
    const next = data.sections.some((s) => s.id === section.id)
      ? data.sections.map((s) => s.id === section.id ? section : s)
      : [...data.sections, section];
    const ok = await saveTab("sections", next);
    setData({ ...data, sections: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Section saved."] : ["warn", "Saved locally."]);
  };
  const deleteSection = async (id) => {
    if (!window.confirm("Delete this section?")) return;
    const next = data.sections.filter((s) => s.id !== id);
    const ok = await saveTab("sections", next);
    setData({ ...data, sections: next });
    setMsg(ok ? ["ok", "Section deleted."] : ["warn", "Deleted locally."]);
  };

  // VARIETIES
  const addVariety = () => {
    setEditForm({ type: "variety", data: { id: "var-" + Date.now(), name: "", commodity: "", description: "" } });
  };
  const saveVariety = async (variety) => {
    const next = data.varieties.some((v) => v.id === variety.id)
      ? data.varieties.map((v) => v.id === variety.id ? variety : v)
      : [...data.varieties, variety];
    const ok = await saveTab("varieties", next);
    setData({ ...data, varieties: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Variety saved."] : ["warn", "Saved locally."]);
  };
  const deleteVariety = async (id) => {
    if (!window.confirm("Delete this variety?")) return;
    const next = data.varieties.filter((v) => v.id !== id);
    const ok = await saveTab("varieties", next);
    setData({ ...data, varieties: next });
    setMsg(ok ? ["ok", "Variety deleted."] : ["warn", "Deleted locally."]);
  };

  // FARM BLOCKS
  const addBlock = () => {
    setEditForm({ type: "block", data: { id: "blk-" + Date.now(), name: "", section: "", area: "", commodity: "", variety: "" } });
  };
  const saveBlock = async (block) => {
    const next = data.blocks.some((b) => b.id === block.id)
      ? data.blocks.map((b) => b.id === block.id ? block : b)
      : [...data.blocks, block];
    const ok = await saveTab("blocks", next);
    setData({ ...data, blocks: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Block saved."] : ["warn", "Saved locally."]);
  };
  const deleteBlock = async (id) => {
    if (!window.confirm("Delete this block?")) return;
    const next = data.blocks.filter((b) => b.id !== id);
    const ok = await saveTab("blocks", next);
    setData({ ...data, blocks: next });
    setMsg(ok ? ["ok", "Block deleted."] : ["warn", "Deleted locally."]);
  };

  // PALLET TYPES
  const addPallet = () => {
    setEditForm({ type: "pallet", data: { id: "pal-" + Date.now(), code: "", name: "", weight: "", capacity: "", dimensions: "", material: "" } });
  };
  const savePallet = async (pallet) => {
    const next = data.pallets.some((p) => p.id === pallet.id)
      ? data.pallets.map((p) => p.id === pallet.id ? pallet : p)
      : [...data.pallets, pallet];
    const ok = await saveTab("pallets", next);
    setData({ ...data, pallets: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Pallet type saved."] : ["warn", "Saved locally."]);
  };
  const deletePallet = async (id) => {
    if (!window.confirm("Delete this pallet type?")) return;
    const next = data.pallets.filter((p) => p.id !== id);
    const ok = await saveTab("pallets", next);
    setData({ ...data, pallets: next });
    setMsg(ok ? ["ok", "Pallet type deleted."] : ["warn", "Deleted locally."]);
  };

  // CRATE TYPES
  const addCrate = () => {
    setEditForm({ type: "crate", data: { id: "crt-" + Date.now(), code: "", name: "", weight: "", capacity: "", dimensions: "", material: "" } });
  };
  const saveCrate = async (crate) => {
    const next = data.crates.some((c) => c.id === crate.id)
      ? data.crates.map((c) => c.id === crate.id ? crate : c)
      : [...data.crates, crate];
    const ok = await saveTab("crates", next);
    setData({ ...data, crates: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Crate type saved."] : ["warn", "Saved locally."]);
  };
  const deleteCrate = async (id) => {
    if (!window.confirm("Delete this crate type?")) return;
    const next = data.crates.filter((c) => c.id !== id);
    const ok = await saveTab("crates", next);
    setData({ ...data, crates: next });
    setMsg(ok ? ["ok", "Crate type deleted."] : ["warn", "Deleted locally."]);
  };

  // GRADES
  const addGrade = () => {
    setEditForm({ type: "grade", data: { id: "grd-" + Date.now(), code: "", name: "", minScore: "", maxScore: "", color: "", description: "" } });
  };
  const saveGrade = async (grade) => {
    const next = data.grades.some((g) => g.id === grade.id)
      ? data.grades.map((g) => g.id === grade.id ? grade : g)
      : [...data.grades, grade];
    const ok = await saveTab("grades", next);
    setData({ ...data, grades: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Grade saved."] : ["warn", "Saved locally."]);
  };
  const deleteGrade = async (id) => {
    if (!window.confirm("Delete this grade?")) return;
    const next = data.grades.filter((g) => g.id !== id);
    const ok = await saveTab("grades", next);
    setData({ ...data, grades: next });
    setMsg(ok ? ["ok", "Grade deleted."] : ["warn", "Deleted locally."]);
  };

  // SEASONS
  const addSeason = () => {
    setEditForm({ type: "season", data: { id: "sea-" + Date.now(), name: "", startMonth: "", endMonth: "", description: "" } });
  };
  const saveSeason = async (season) => {
    const next = data.seasons.some((s) => s.id === season.id)
      ? data.seasons.map((s) => s.id === season.id ? season : s)
      : [...data.seasons, season];
    const ok = await saveTab("seasons", next);
    setData({ ...data, seasons: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Season saved."] : ["warn", "Saved locally."]);
  };
  const deleteSeason = async (id) => {
    if (!window.confirm("Delete this season?")) return;
    const next = data.seasons.filter((s) => s.id !== id);
    const ok = await saveTab("seasons", next);
    setData({ ...data, seasons: next });
    setMsg(ok ? ["ok", "Season deleted."] : ["warn", "Deleted locally."]);
  };

  // CHANNELS
  const addChannel = () => {
    setEditForm({ type: "channel", data: { id: "ch-" + Date.now(), name: "", type: "", description: "" } });
  };
  const saveChannel = async (channel) => {
    const next = data.channels.some((c) => c.id === channel.id)
      ? data.channels.map((c) => c.id === channel.id ? channel : c)
      : [...data.channels, channel];
    const ok = await saveTab("channels", next);
    setData({ ...data, channels: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Channel saved."] : ["warn", "Saved locally."]);
  };
  const deleteChannel = async (id) => {
    if (!window.confirm("Delete this channel?")) return;
    const next = data.channels.filter((c) => c.id !== id);
    const ok = await saveTab("channels", next);
    setData({ ...data, channels: next });
    setMsg(ok ? ["ok", "Channel deleted."] : ["warn", "Deleted locally."]);
  };

  // TRANSPORT
  const addTransport = () => {
    setEditForm({ type: "transport", data: { id: "tr-" + Date.now(), code: "", name: "", type: "", capacity: "", description: "" } });
  };
  const saveTransport = async (transport) => {
    const next = data.transport.some((t) => t.id === transport.id)
      ? data.transport.map((t) => t.id === transport.id ? transport : t)
      : [...data.transport, transport];
    const ok = await saveTab("transport", next);
    setData({ ...data, transport: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Transport saved."] : ["warn", "Saved locally."]);
  };
  const deleteTransport = async (id) => {
    if (!window.confirm("Delete this transport?")) return;
    const next = data.transport.filter((t) => t.id !== id);
    const ok = await saveTab("transport", next);
    setData({ ...data, transport: next });
    setMsg(ok ? ["ok", "Transport deleted."] : ["warn", "Deleted locally."]);
  };

  // PACK TYPES
  const addPackType = () => {
    setEditForm({ type: "packtype", data: { id: "pt-" + Date.now(), code: "", name: "", description: "" } });
  };
  const savePackType = async (packtype) => {
    const next = data.packtypes.some((p) => p.id === packtype.id)
      ? data.packtypes.map((p) => p.id === packtype.id ? packtype : p)
      : [...data.packtypes, packtype];
    const ok = await saveTab("packtypes", next);
    setData({ ...data, packtypes: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Pack type saved."] : ["warn", "Saved locally."]);
  };
  const deletePackType = async (id) => {
    if (!window.confirm("Delete this pack type?")) return;
    const next = data.packtypes.filter((p) => p.id !== id);
    const ok = await saveTab("packtypes", next);
    setData({ ...data, packtypes: next });
    setMsg(ok ? ["ok", "Pack type deleted."] : ["warn", "Deleted locally."]);
  };

  // SIZES
  const addSize = () => {
    setEditForm({ type: "size", data: { id: "sz-" + Date.now(), code: "", name: "", count: "", description: "" } });
  };
  const saveSize = async (size) => {
    const next = data.sizes.some((s) => s.id === size.id)
      ? data.sizes.map((s) => s.id === size.id ? size : s)
      : [...data.sizes, size];
    const ok = await saveTab("sizes", next);
    setData({ ...data, sizes: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Size saved."] : ["warn", "Saved locally."]);
  };
  const deleteSize = async (id) => {
    if (!window.confirm("Delete this size?")) return;
    const next = data.sizes.filter((s) => s.id !== id);
    const ok = await saveTab("sizes", next);
    setData({ ...data, sizes: next });
    setMsg(ok ? ["ok", "Size deleted."] : ["warn", "Deleted locally."]);
  };

  // BRANDS
  const addBrand = () => {
    setEditForm({ type: "brand", data: { id: "br-" + Date.now(), name: "", description: "" } });
  };
  const saveBrand = async (brand) => {
    const next = data.brands.some((b) => b.id === brand.id)
      ? data.brands.map((b) => b.id === brand.id ? brand : b)
      : [...data.brands, brand];
    const ok = await saveTab("brands", next);
    setData({ ...data, brands: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Brand saved."] : ["warn", "Saved locally."]);
  };
  const deleteBrand = async (id) => {
    if (!window.confirm("Delete this brand?")) return;
    const next = data.brands.filter((b) => b.id !== id);
    const ok = await saveTab("brands", next);
    setData({ ...data, brands: next });
    setMsg(ok ? ["ok", "Brand deleted."] : ["warn", "Deleted locally."]);
  };

  // INVENTORY CODES
  const addInvCode = () => {
    setEditForm({ type: "invcode", data: { id: "ic-" + Date.now(), code: "", commodity: "", variety: "", brand: "", description: "" } });
  };
  const saveInvCode = async (invcode) => {
    const next = data.invcodes.some((i) => i.id === invcode.id)
      ? data.invcodes.map((i) => i.id === invcode.id ? invcode : i)
      : [...data.invcodes, invcode];
    const ok = await saveTab("invcodes", next);
    setData({ ...data, invcodes: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Inventory code saved."] : ["warn", "Saved locally."]);
  };
  const deleteInvCode = async (id) => {
    if (!window.confirm("Delete this inventory code?")) return;
    const next = data.invcodes.filter((i) => i.id !== id);
    const ok = await saveTab("invcodes", next);
    setData({ ...data, invcodes: next });
    setMsg(ok ? ["ok", "Inventory code deleted."] : ["warn", "Deleted locally."]);
  };

  // CUSTOMERS
  const addCustomer = () => {
    setEditForm({ type: "customer", data: { id: "cust-" + Date.now(), name: "", contactPerson: "", phone: "", email: "", country: "", notes: "" } });
  };
  const saveCustomer = async (customer) => {
    const next = data.customers.some((c) => c.id === customer.id)
      ? data.customers.map((c) => c.id === customer.id ? customer : c)
      : [...data.customers, customer];
    const ok = await saveTab("customers", next);
    setData({ ...data, customers: next });
    setEditForm(null);
    setMsg(ok ? ["ok", "Customer saved."] : ["warn", "Saved locally."]);
  };
  const deleteCustomer = async (id) => {
    if (!window.confirm("Delete this customer?")) return;
    const next = data.customers.filter((c) => c.id !== id);
    const ok = await saveTab("customers", next);
    setData({ ...data, customers: next });
    setMsg(ok ? ["ok", "Customer deleted."] : ["warn", "Deleted locally."]);
  };

  // Form Component
  const EditForm = ({ form, onSave, onCancel, data }) => {
    const [item, setItem] = useState(form.data);
    const set = (k) => (e) => setItem({ ...item, [k]: e.target.value });

    if (form.type === "section") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="e.g., Bottom, Middle, Top" /></label>
            <label className="fld"><span>Area (hectares)</span><input className="inp" type="number" value={item.area || ""} onChange={set("area")} placeholder="0.00" step="0.1" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Optional details about this section" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "commodity") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Commodity name" /></label>
            <label className="fld"><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Optional" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "variety") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Variety name" /></label>
            <label className="fld"><span>Commodity *</span><select className="sel" value={item.commodity || ""} onChange={set("commodity")}><option value="">-- Select commodity --</option>{data?.commodities?.map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}</select></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Optional" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "block") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Block Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="e.g., C09" /></label>
            <label className="fld"><span>Section *</span><select className="sel" value={item.section || ""} onChange={set("section")}><option value="">-- Select section --</option>{data?.sections?.map((s) => <option key={s.id} value={s.name}>{s.name}</option>)}</select></label>
            <label className="fld"><span>Area (hectares)</span><input className="inp" type="number" value={item.area || ""} onChange={set("area")} placeholder="0.00" step="0.1" /></label>
            <label className="fld"><span>Commodity</span><select className="sel" value={item.commodity || ""} onChange={set("commodity")}><option value="">-- Select commodity --</option>{data?.commodities?.map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}</select></label>
            <label className="fld"><span>Variety</span><select className="sel" value={item.variety || ""} onChange={set("variety")}><option value="">-- Select variety --</option>{data?.varieties?.map((v) => <option key={v.id} value={v.name}>{v.name}</option>)}</select></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "pallet") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Code *</span><input className="inp" value={item.code} onChange={set("code")} placeholder="e.g., STD-PAL" /></label>
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Pallet type name" /></label>
            <label className="fld"><span>Weight (kg)</span><input className="inp" type="number" value={item.weight || ""} onChange={set("weight")} placeholder="0.00" step="0.1" /></label>
            <label className="fld"><span>Capacity (kg)</span><input className="inp" type="number" value={item.capacity || ""} onChange={set("capacity")} placeholder="0" step="1" /></label>
            <label className="fld"><span>Dimensions</span><input className="inp" value={item.dimensions || ""} onChange={set("dimensions")} placeholder="L×W×H cm" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Material</span><input className="inp" value={item.material || ""} onChange={set("material")} placeholder="e.g., Wood, Plastic" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "crate") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Code *</span><input className="inp" value={item.code} onChange={set("code")} placeholder="e.g., CRT-L" /></label>
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Crate type name" /></label>
            <label className="fld"><span>Weight (kg)</span><input className="inp" type="number" value={item.weight || ""} onChange={set("weight")} placeholder="0.00" step="0.1" /></label>
            <label className="fld"><span>Capacity (kg)</span><input className="inp" type="number" value={item.capacity || ""} onChange={set("capacity")} placeholder="0" step="0.1" /></label>
            <label className="fld"><span>Dimensions</span><input className="inp" value={item.dimensions || ""} onChange={set("dimensions")} placeholder="L×W×H cm" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Material</span><input className="inp" value={item.material || ""} onChange={set("material")} placeholder="e.g., Cardboard, Plastic" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "grade") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Code *</span><input className="inp" value={item.code} onChange={set("code")} placeholder="e.g., A, B, C" /></label>
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Grade name" /></label>
            <label className="fld"><span>Min Score (%)</span><input className="inp" type="number" value={item.minScore || ""} onChange={set("minScore")} placeholder="0" min="0" max="100" /></label>
            <label className="fld"><span>Max Score (%)</span><input className="inp" type="number" value={item.maxScore || ""} onChange={set("maxScore")} placeholder="100" min="0" max="100" /></label>
            <label className="fld"><span>Color (Badge)</span><input className="inp" value={item.color || ""} onChange={set("color")} placeholder="e.g., #2F7D53" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Quality requirements" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "season") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="e.g., Early Season" /></label>
            <label className="fld"><span>Start Month</span><input className="inp" value={item.startMonth || ""} onChange={set("startMonth")} placeholder="January" /></label>
            <label className="fld"><span>End Month</span><input className="inp" value={item.endMonth || ""} onChange={set("endMonth")} placeholder="March" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Season details" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "channel") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="e.g., Export, Domestic" /></label>
            <label className="fld"><span>Type</span><input className="inp" value={item.type || ""} onChange={set("type")} placeholder="e.g., Direct, Distributor" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Channel details" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "transport") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Code *</span><input className="inp" value={item.code} onChange={set("code")} placeholder="e.g., TRUCK1" /></label>
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Transport name" /></label>
            <label className="fld"><span>Type</span><input className="inp" value={item.type || ""} onChange={set("type")} placeholder="e.g., Truck, Container, Air" /></label>
            <label className="fld"><span>Capacity</span><input className="inp" value={item.capacity || ""} onChange={set("capacity")} placeholder="e.g., 5 tons" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Transport details" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "packtype") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Code *</span><input className="inp" value={item.code} onChange={set("code")} placeholder="e.g., BOX" /></label>
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Pack type name" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Packing method details" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "size") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Code *</span><input className="inp" value={item.code} onChange={set("code")} placeholder="e.g., S1, S2" /></label>
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Size name" /></label>
            <label className="fld"><span>Count/Pieces</span><input className="inp" value={item.count || ""} onChange={set("count")} placeholder="e.g., 24, 48" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Size details" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "brand") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Brand name" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Brand information" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "invcode") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Code *</span><input className="inp" value={item.code} onChange={set("code")} placeholder="Unique SKU/code" /></label>
            <label className="fld"><span>Commodity</span><input className="inp" value={item.commodity || ""} onChange={set("commodity")} placeholder="e.g., Blueberry" /></label>
            <label className="fld"><span>Variety</span><input className="inp" value={item.variety || ""} onChange={set("variety")} placeholder="e.g., ARANA" /></label>
            <label className="fld"><span>Brand</span><input className="inp" value={item.brand || ""} onChange={set("brand")} placeholder="Brand name" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Description</span><input className="inp" value={item.description || ""} onChange={set("description")} placeholder="Product details" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    if (form.type === "customer") {
      return (
        <Card>
          <div className="fgrid">
            <label className="fld"><span>Name *</span><input className="inp" value={item.name} onChange={set("name")} placeholder="Customer name" /></label>
            <label className="fld"><span>Contact Person</span><input className="inp" value={item.contactPerson || ""} onChange={set("contactPerson")} placeholder="Name" /></label>
            <label className="fld"><span>Phone</span><input className="inp" value={item.phone || ""} onChange={set("phone")} placeholder="+1-XXX-XXX-XXXX" /></label>
            <label className="fld"><span>Email</span><input className="inp" type="email" value={item.email || ""} onChange={set("email")} placeholder="email@example.com" /></label>
            <label className="fld"><span>Country</span><input className="inp" value={item.country || ""} onChange={set("country")} placeholder="Country name" /></label>
            <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Notes</span><input className="inp" value={item.notes || ""} onChange={set("notes")} placeholder="Additional information" /></label>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={() => onSave(item)}>Save</button>
            <button className="btn btn-mini" onClick={onCancel}>Cancel</button>
          </div>
        </Card>
      );
    }

    return null;
  };

  if (editForm) {
    return (
      <>
        <div className="page-head">
          <h2>Edit {editForm.type}</h2>
          <button className="btn btn-mini" onClick={() => setEditForm(null)}>Back</button>
        </div>
        {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom: 12 }}>{msg[1]}</p>}
        <EditForm
          form={editForm}
          data={data}
          onSave={
            editForm.type === "section" ? saveSection :
            editForm.type === "commodity" ? saveCommodity :
            editForm.type === "variety" ? saveVariety :
            editForm.type === "block" ? saveBlock :
            editForm.type === "pallet" ? savePallet :
            editForm.type === "crate" ? saveCrate :
            editForm.type === "grade" ? saveGrade :
            editForm.type === "season" ? saveSeason :
            editForm.type === "channel" ? saveChannel :
            editForm.type === "transport" ? saveTransport :
            editForm.type === "packtype" ? savePackType :
            editForm.type === "size" ? saveSize :
            editForm.type === "brand" ? saveBrand :
            editForm.type === "invcode" ? saveInvCode :
            saveCustomer
          }
          onCancel={() => setEditForm(null)}
        />
      </>
    );
  }

  const tabs = [
    { key: "sections", label: "Sections", count: data.sections.length, onAdd: addSection, data: data.sections, onDelete: deleteSection },
    { key: "commodities", label: "Commodities", count: data.commodities.length, onAdd: addCommodity, data: data.commodities, onDelete: deleteCommodity },
    { key: "varieties", label: "Varieties", count: data.varieties.length, onAdd: addVariety, data: data.varieties, onDelete: deleteVariety },
    { key: "blocks", label: "Farm Blocks", count: data.blocks.length, onAdd: addBlock, data: data.blocks, onDelete: deleteBlock },
    { key: "pallets", label: "Pallet Types", count: data.pallets.length, onAdd: addPallet, data: data.pallets, onDelete: deletePallet },
    { key: "crates", label: "Crate Types", count: data.crates.length, onAdd: addCrate, data: data.crates, onDelete: deleteCrate },
    { key: "grades", label: "Grades", count: data.grades.length, onAdd: addGrade, data: data.grades, onDelete: deleteGrade },
    { key: "seasons", label: "Seasons", count: data.seasons.length, onAdd: addSeason, data: data.seasons, onDelete: deleteSeason },
    { key: "channels", label: "Channels", count: data.channels.length, onAdd: addChannel, data: data.channels, onDelete: deleteChannel },
    { key: "transport", label: "Transport", count: data.transport.length, onAdd: addTransport, data: data.transport, onDelete: deleteTransport },
    { key: "packtypes", label: "Pack Types", count: data.packtypes.length, onAdd: addPackType, data: data.packtypes, onDelete: deletePackType },
    { key: "sizes", label: "Size/Count", count: data.sizes.length, onAdd: addSize, data: data.sizes, onDelete: deleteSize },
    { key: "brands", label: "Brands", count: data.brands.length, onAdd: addBrand, data: data.brands, onDelete: deleteBrand },
    { key: "invcodes", label: "Inv Codes", count: data.invcodes.length, onAdd: addInvCode, data: data.invcodes, onDelete: deleteInvCode },
    { key: "customers", label: "Customers", count: data.customers.length, onAdd: addCustomer, data: data.customers, onDelete: deleteCustomer },
  ];

  const currentTab = tabs.find((t) => t.key === tab);

  if (!isAdmin) {
    return (
      <>
        <div className="page-head">
          <h2>Master Files</h2>
        </div>
        <Card>
          <p style={{ padding: 24, textAlign: "center", color: "var(--red)" }}>
            ⛔ Master Files access is restricted to Administrators only.
          </p>
        </Card>
      </>
    );
  }

  return (
    <>
      <div className="page-head">
        <h2>Master Files</h2>
        <button className="btn btn-primary" onClick={currentTab?.onAdd}>+ Add {currentTab?.label}</button>
      </div>
      {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom: 12 }}>{msg[1]}</p>}

      <div className="tabs">
        {tabs.map((t) => (
          <button key={t.key} className={"tab" + (tab === t.key ? " active" : "")} onClick={() => setTab(t.key)}>
            {t.label} <span style={{ marginLeft: 6, opacity: 0.7 }}>({t.count})</span>
          </button>
        ))}
      </div>

      <Card title={currentTab?.label} pad={false}>
        {loading ? (
          <p className="muted" style={{ padding: 24 }}>Loading…</p>
        ) : currentTab?.data?.length === 0 ? (
          <p className="muted" style={{ padding: 24 }}>No {currentTab?.label?.toLowerCase()} yet. Click the add button to create one.</p>
        ) : tab === "sections" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Name</th><th>Area (ha)</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "section", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td>{item.area || "—"}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteSection(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "commodities" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Name</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "commodity", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteCommodity(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "varieties" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Name</th><th>Commodity</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "variety", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td>{item.commodity || "—"}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteVariety(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "blocks" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Block</th><th>Section</th><th>Area (ha)</th><th>Commodity</th><th>Variety</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "block", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td>{item.section || "—"}</td>
                    <td className="r">{item.area || "—"}</td>
                    <td>{item.commodity || "—"}</td>
                    <td>{item.variety || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteBlock(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "pallets" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Code</th><th>Name</th><th>Capacity (kg)</th><th>Dimensions</th><th>Material</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "pallet", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.code}</strong></td>
                    <td>{item.name}</td>
                    <td className="r">{item.capacity || "—"}</td>
                    <td className="muted">{item.dimensions || "—"}</td>
                    <td>{item.material || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deletePallet(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "crates" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Code</th><th>Name</th><th>Capacity (kg)</th><th>Dimensions</th><th>Material</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "crate", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.code}</strong></td>
                    <td>{item.name}</td>
                    <td className="r">{item.capacity || "—"}</td>
                    <td className="muted">{item.dimensions || "—"}</td>
                    <td>{item.material || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteCrate(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "grades" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Code</th><th>Name</th><th>Score Range</th><th>Color</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "grade", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.code}</strong></td>
                    <td>{item.name}</td>
                    <td>{item.minScore || "0"}% - {item.maxScore || "100"}%</td>
                    <td><span style={{ display: "inline-block", width: 24, height: 24, borderRadius: 4, background: item.color || "#ccc", border: "1px solid #999" }} title={item.color} /></td>
                    <td className="muted" style={{ fontSize: 11 }}>{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteGrade(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "seasons" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Name</th><th>Start Month</th><th>End Month</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "season", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td>{item.startMonth || "—"}</td>
                    <td>{item.endMonth || "—"}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteSeason(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "channels" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Name</th><th>Type</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "channel", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td>{item.type || "—"}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteChannel(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "transport" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Code</th><th>Name</th><th>Type</th><th>Capacity</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "transport", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.code}</strong></td>
                    <td>{item.name}</td>
                    <td>{item.type || "—"}</td>
                    <td>{item.capacity || "—"}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteTransport(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "packtypes" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Code</th><th>Name</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "packtype", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.code}</strong></td>
                    <td>{item.name}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deletePackType(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "sizes" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Code</th><th>Name</th><th>Count/Pieces</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "size", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.code}</strong></td>
                    <td>{item.name}</td>
                    <td className="r">{item.count || "—"}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteSize(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "brands" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Name</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "brand", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteBrand(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "invcodes" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Code</th><th>Commodity</th><th>Variety</th><th>Brand</th><th>Description</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "invcode", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.code}</strong></td>
                    <td>{item.commodity || "—"}</td>
                    <td>{item.variety || "—"}</td>
                    <td>{item.brand || "—"}</td>
                    <td className="muted">{item.description || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteInvCode(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : tab === "customers" ? (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr><th>Name</th><th>Contact Person</th><th>Phone</th><th>Email</th><th>Country</th><th></th></tr>
              </thead>
              <tbody>
                {currentTab.data.map((item) => (
                  <tr key={item.id} onClick={() => setEditForm({ type: "customer", data: item })} style={{ cursor: "pointer" }}>
                    <td><strong>{item.name}</strong></td>
                    <td>{item.contactPerson || "—"}</td>
                    <td className="muted">{item.phone || "—"}</td>
                    <td className="muted">{item.email || "—"}</td>
                    <td>{item.country || "—"}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini btn-danger" onClick={(e) => { e.stopPropagation(); deleteCustomer(item.id); }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : null}
      </Card>
    </>
  );
}

/* ================= FRUIT RECEIVING ================= */
const FR_KEY = "ml-fruit-receiving-v1";

function FruitReceiving({ me }) {
  const [records, setRecords] = useState([]);
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState(null);
  const [msg, setMsg] = useState(null);
  const [filter, setFilter] = useState("all");
  const [searchDate, setSearchDate] = useState(new Date().toISOString().slice(0, 10));
  const [masterData, setMasterData] = useState({
    sections: [],
    seasons: [],
    varieties: [],
    crates: [],
    blocks: [],
    pallets: [],
    grades: [],
  });
  const [printLabel, setPrintLabel] = useState(null);

  useEffect(() => {
    const loadData = async () => {
      try {
        // Load batch records
        const records = await loadShared(FR_KEY);
        setRecords(Array.isArray(records) ? records.filter((x) => x && x.id) : []);
        
        // Load master data (parse JSON strings)
        const sectionsData = await loadShared(MF_KEYS.sections);
        const seasonsData = await loadShared(MF_KEYS.seasons);
        const varietiesData = await loadShared(MF_KEYS.varieties);
        const cratesData = await loadShared(MF_KEYS.crates);
        const blocksData = await loadShared(MF_KEYS.blocks);
        const palletsData = await loadShared(MF_KEYS.pallets);
        const gradesData = await loadShared(MF_KEYS.grades);
        
        // Parse JSON if string, ensure array
        const parseMasterData = (data) => {
          if (!data) return [];
          if (typeof data === 'string') {
            try {
              const parsed = JSON.parse(data);
              return Array.isArray(parsed) ? parsed : [];
            } catch (e) {
              console.error("JSON parse error:", e);
              return [];
            }
          }
          return Array.isArray(data) ? data : [];
        };
        
        setMasterData({
          sections: parseMasterData(sectionsData),
          seasons: parseMasterData(seasonsData),
          varieties: parseMasterData(varietiesData),
          crates: parseMasterData(cratesData),
          blocks: parseMasterData(blocksData),
          pallets: parseMasterData(palletsData),
          grades: parseMasterData(gradesData),
        });
      } catch (e) {
        console.error("Load error:", e);
      }
      setLoading(false);
    };
    loadData();
  }, []);

  const calculateNetWeight = () => {
    if (!form.lugQuantity || !form.lugType || !form.palletType || !form.grossWeight) {
      return "";
    }
    
    // Helper to extract weight from master data items
    const getWeight = (item) => {
      if (!item) return 0;
      if (typeof item === 'string') return 0;
      if (typeof item !== 'object') return 0;
      
      // Try multiple property names: weight, wt, Weight
      const weight = item.weight || item.wt || item.Weight;
      const parsed = parseFloat(weight);
      return isNaN(parsed) ? 0 : parsed;
    };
    
    // Find crate by name
    const crate = masterData.crates.find((c) => 
      (typeof c === 'string' ? c : (c.name || c)) === form.lugType
    );
    const crateWeight = getWeight(crate);
    
    // Find pallet by name
    const pallet = masterData.pallets.find((p) => 
      (typeof p === 'string' ? p : (p.name || p)) === form.palletType
    );
    const palletWeight = getWeight(pallet);
    
    // Calculate net weight: Gross - (Qty × Crate Wt + Pallet Wt)
    const lugQty = parseFloat(form.lugQuantity) || 0;
    const grossWt = parseFloat(form.grossWeight) || 0;
    const packagingWeight = (lugQty * crateWeight) + palletWeight;
    const netWt = grossWt - packagingWeight;
    
    // Debug logging (can be removed after testing)
    if (process.env.NODE_ENV === 'development' || true) {
      console.log('NET WEIGHT CALC DEBUG:', {
        lugQty,
        crateWeight,
        palletWeight,
        packagingWeight,
        grossWt,
        netWt,
        crateName: form.lugType,
        palletName: form.palletType
      });
    }
    
    return netWt >= 0 ? netWt.toFixed(2) : "0.00";
  };

  const newRecord = () => {
    setForm({
      id: "fr-" + Date.now(),
      harvestDate: new Date().toISOString().slice(0, 10),
      season: "",
      section: "",
      variety: "",
      block: "",
      grade: "",
      batchNumber: "",
      supplier: "Merrylight",
      lugType: "",
      lugQuantity: "",
      palletType: "",
      grossWeight: "",
      netQuantity: "",
      receivedBy: me.name,
      timestamp: new Date().toISOString(),
    });
  };

  const saveRecord = async () => {
    if (!form.harvestDate || !form.supplier || !form.variety || !form.section || !form.batchNumber) {
      setMsg(["warn", "Please fill in all required fields."]);
      return;
    }
    const next = records.some((r) => r.id === form.id)
      ? records.map((r) => r.id === form.id ? form : r)
      : [...records, form];
    const ok = await saveShared(FR_KEY, next);
    setRecords(next);
    if (ok) {
      setPrintLabel(form);
      setMsg(["ok", "Entry saved. Click 'Print Label' to print batch label."]);
      setTimeout(() => setForm(null), 2000);
    } else {
      setMsg(["warn", "Saved locally but shared storage unavailable."]);
    }
  };

  const deleteRecord = async (id) => {
    if (!window.confirm("Delete this receiving record?")) return;
    const next = records.filter((r) => r.id !== id);
    const ok = await saveShared(FR_KEY, next);
    setRecords(next);
    setMsg(ok ? ["ok", "Record deleted."] : ["warn", "Deleted locally."]);
  };

  const filtered = records.filter((r) => {
    if (r.harvestDate && r.harvestDate < searchDate) return false;
    return true;
  }).sort((a, b) => new Date(b.harvestDate || 0) - new Date(a.harvestDate || 0));

  const totalQty = filtered.reduce((sum, r) => sum + (num(r.netQuantity) || 0), 0);
  const varieties = [...new Set(filtered.map((r) => r.variety).filter(Boolean))].length;

  if (printLabel) {
    // Generate real barcode using JsBarcode library (after component renders)
    if (typeof window !== 'undefined') {
      setTimeout(() => {
        // Load JsBarcode library if not already loaded
        if (!window.JsBarcode) {
          const script = document.createElement('script');
          script.src = 'https://cdn.jsdelivr.net/npm/jsbarcode@3.11.5/dist/JsBarcode.all.min.js';
          script.async = true;
          script.onload = () => {
            if (window.JsBarcode && printLabel?.batchNumber) {
              try {
                window.JsBarcode("#barcode-canvas", printLabel.batchNumber, {
                  format: "CODE128",
                  width: 2,
                  height: 40,
                  displayValue: false
                });
              } catch (e) {
                console.error("Barcode generation error:", e);
              }
            }
          };
          document.body.appendChild(script);
        } else if (window.JsBarcode && printLabel?.batchNumber) {
          // Library already loaded, generate barcode
          try {
            window.JsBarcode("#barcode-canvas", printLabel.batchNumber, {
              format: "CODE128",
              width: 2,
              height: 40,
              displayValue: false
            });
          } catch (e) {
            console.error("Barcode generation error:", e);
          }
        }
      }, 100);
    }

    return (
      <div style={{ padding: 20 }}>
        <button className="btn btn-mini" onClick={() => { setPrintLabel(null); window.print(); }} style={{ marginBottom: 16 }}>Done Printing</button>
        
        {/* LABEL PRINT TEMPLATE - 100mm x 70mm */}
        <div style={{
          width: "100mm",
          height: "70mm",
          border: "1px solid #000",
          padding: "4mm",
          fontFamily: "Arial, sans-serif",
          fontSize: "10px",
          margin: "20px auto",
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-between",
          boxSizing: "border-box",
          backgroundColor: "#fff",
          pageBreakAfter: "always"
        }}>
          {/* Batch Number and Barcode */}
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: "16px", fontWeight: "bold", marginBottom: "2mm" }}>{printLabel.batchNumber}</div>
            <div style={{ marginBottom: "2mm", minHeight: "50px", display: "flex", justifyContent: "center", alignItems: "center" }}>
              <canvas id="barcode-canvas" style={{ maxWidth: "100%", height: "auto", margin: "0 auto", display: "block" }}></canvas>
            </div>
          </div>

          {/* Details Grid */}
          <table style={{ width: "100%", fontSize: "11px", borderCollapse: "collapse" }}>
            <tbody>
              <tr>
                <td style={{ paddingRight: "4mm", fontWeight: "bold" }}>Variety:</td>
                <td style={{ fontWeight: "bold" }}>{printLabel.variety}</td>
              </tr>
              <tr>
                <td style={{ paddingRight: "4mm", fontWeight: "bold" }}>Farm Block:</td>
                <td style={{ fontWeight: "bold" }}>{printLabel.block}</td>
              </tr>
              <tr>
                <td style={{ paddingRight: "4mm", fontWeight: "bold" }}>Units:</td>
                <td style={{ fontWeight: "bold" }}>{printLabel.lugQuantity}</td>
              </tr>
              <tr>
                <td style={{ paddingRight: "4mm", fontWeight: "bold" }}>Net Weight:</td>
                <td style={{ fontWeight: "bold" }}>{printLabel.netQuantity} kg</td>
              </tr>
              <tr>
                <td style={{ paddingRight: "4mm", fontWeight: "bold" }}>Harvest Date:</td>
                <td style={{ fontWeight: "bold" }}>{printLabel.harvestDate}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  if (form) {
    return (
      <>
        <div className="page-head">
          <h2>Batch Upload</h2>
          <button className="btn btn-mini" onClick={() => setForm(null)}>Back to list</button>
        </div>
        <Card>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 16 }}>
            {/* LEFT COLUMN */}
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <label className="fld"><span>Harvest Date *</span><input className="inp" type="date" value={form.harvestDate} onChange={(e) => setForm({ ...form, harvestDate: e.target.value })} /></label>
              <label className="fld"><span>Section *</span><select className="sel" value={form.section} onChange={(e) => setForm({ ...form, section: e.target.value })}><option value="">-- Select section --</option>{masterData.sections.length === 0 ? <option disabled>No sections available</option> : masterData.sections.map((s, idx) => <option key={s.id || idx} value={typeof s === 'string' ? s : (s.name || s)}>{typeof s === 'string' ? s : (s.name || s)}</option>)}</select></label>
              <label className="fld"><span>Block *</span><select className="sel" value={form.block} onChange={(e) => setForm({ ...form, block: e.target.value })} disabled={!form.section}><option value="">-- Select block --</option>{form.section && masterData.blocks.length > 0 ? masterData.blocks.filter((b) => (typeof b === 'string' ? b : (b.section || b)) === form.section).length === 0 ? <option disabled>No blocks for this section</option> : masterData.blocks.filter((b) => (typeof b === 'string' ? b : (b.section || b)) === form.section).map((b, idx) => <option key={b.id || idx} value={typeof b === 'string' ? b : (b.name || b)}>{typeof b === 'string' ? b : (b.name || b)}</option>) : <option disabled>Select section first</option>}</select></label>
              <label className="fld"><span>Grade</span><select className="sel" value={form.grade} onChange={(e) => setForm({ ...form, grade: e.target.value })}><option value="">-- Select grade --</option>{masterData.grades.length === 0 ? <option disabled>No grades available</option> : masterData.grades.map((g, idx) => <option key={g.id || idx} value={typeof g === 'string' ? g : (g.name || g)}>{typeof g === 'string' ? g : (g.name || g)}</option>)}</select></label>
              <label className="fld"><span>Supplier *</span><input className="inp" value={form.supplier} onChange={(e) => setForm({ ...form, supplier: e.target.value })} placeholder="Supplier name" /></label>
              <label className="fld"><span>Lug Quantity</span><input className="inp" type="number" value={form.lugQuantity} onChange={(e) => { const newForm = { ...form, lugQuantity: e.target.value }; if (newForm.lugQuantity && newForm.lugType && newForm.palletType && newForm.grossWeight) { const crate = masterData.crates.find((c) => (typeof c === 'string' ? c : (c.name || c)) === newForm.lugType); const crateWt = crate ? parseFloat(crate.weight || crate.wt || crate.Weight || 0) : 0; const pallet = masterData.pallets.find((p) => (typeof p === 'string' ? p : (p.name || p)) === newForm.palletType); const palletWt = pallet ? parseFloat(pallet.weight || pallet.wt || pallet.Weight || 0) : 0; const lugQty = parseFloat(newForm.lugQuantity) || 0; const grossWt = parseFloat(newForm.grossWeight) || 0; const packagingWeight = (lugQty * crateWt) + palletWt; newForm.netQuantity = (grossWt - packagingWeight >= 0 ? (grossWt - packagingWeight).toFixed(2) : "0.00"); } setForm(newForm); }} placeholder="0" /></label>
              <label className="fld"><span>Gross Weight (kg)</span><input className="inp" type="number" value={form.grossWeight} onChange={(e) => { const newForm = { ...form, grossWeight: e.target.value }; if (newForm.lugQuantity && newForm.lugType && newForm.palletType && newForm.grossWeight) { const crate = masterData.crates.find((c) => (typeof c === 'string' ? c : (c.name || c)) === newForm.lugType); const crateWt = crate ? parseFloat(crate.weight || crate.wt || crate.Weight || 0) : 0; const pallet = masterData.pallets.find((p) => (typeof p === 'string' ? p : (p.name || p)) === newForm.palletType); const palletWt = pallet ? parseFloat(pallet.weight || pallet.wt || pallet.Weight || 0) : 0; const lugQty = parseFloat(newForm.lugQuantity) || 0; const grossWt = parseFloat(newForm.grossWeight) || 0; const packagingWeight = (lugQty * crateWt) + palletWt; newForm.netQuantity = (grossWt - packagingWeight >= 0 ? (grossWt - packagingWeight).toFixed(2) : "0.00"); } setForm(newForm); }} placeholder="0" step="0.1" /></label>
            </div>

            {/* RIGHT COLUMN */}
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <label className="fld"><span>Season *</span><select className="sel" value={form.season} onChange={(e) => { const newSeason = e.target.value; const newForm = { ...form, season: newSeason }; if (newSeason) { const seasonParts = newSeason.split('/'); const seasonSuffix = seasonParts[seasonParts.length - 1]; const batchesForSeason = records.filter((r) => r.batchNumber && r.batchNumber.startsWith('MER' + seasonSuffix)); let counter = 1; if (batchesForSeason.length > 0) { const sorted = batchesForSeason.sort((a, b) => parseInt(b.batchNumber.slice(-5)) - parseInt(a.batchNumber.slice(-5))); const lastNum = parseInt(sorted[0].batchNumber.slice(-5)); counter = lastNum + 1; } newForm.batchNumber = `MER${seasonSuffix}${String(counter).padStart(5, '0')}`; } setForm(newForm); }}><option value="">-- Select season --</option>{masterData.seasons.length === 0 ? <option disabled>No seasons available</option> : masterData.seasons.map((s, idx) => <option key={s.id || idx} value={typeof s === 'string' ? s : (s.name || s)}>{typeof s === 'string' ? s : (s.name || s)}</option>)}</select></label>
              <label className="fld"><span>Variety *</span><select className="sel" value={form.variety} onChange={(e) => setForm({ ...form, variety: e.target.value })}><option value="">-- Select variety --</option>{masterData.varieties.length === 0 ? <option disabled>No varieties available</option> : masterData.varieties.map((v, idx) => <option key={v.id || idx} value={typeof v === 'string' ? v : (v.name || v)}>{typeof v === 'string' ? v : (v.name || v)}</option>)}</select></label>
              <label className="fld"><span>Batch Number *</span><input className="inp" value={form.batchNumber} onChange={(e) => setForm({ ...form, batchNumber: e.target.value })} placeholder="Auto-generated or enter" /></label>
              <label className="fld"><span>Lug Type</span><select className="sel" value={form.lugType} onChange={(e) => { const newForm = { ...form, lugType: e.target.value }; if (newForm.lugQuantity && newForm.lugType && newForm.palletType && newForm.grossWeight) { const crate = masterData.crates.find((c) => (typeof c === 'string' ? c : (c.name || c)) === newForm.lugType); const crateWt = crate ? parseFloat(crate.weight || crate.wt || crate.Weight || 0) : 0; const pallet = masterData.pallets.find((p) => (typeof p === 'string' ? p : (p.name || p)) === newForm.palletType); const palletWt = pallet ? parseFloat(pallet.weight || pallet.wt || pallet.Weight || 0) : 0; const lugQty = parseFloat(newForm.lugQuantity) || 0; const grossWt = parseFloat(newForm.grossWeight) || 0; const packagingWeight = (lugQty * crateWt) + palletWt; newForm.netQuantity = (grossWt - packagingWeight >= 0 ? (grossWt - packagingWeight).toFixed(2) : "0.00"); } setForm(newForm); }}><option value="">-- Select lug type --</option>{masterData.crates.length === 0 ? <option disabled>No lug types available</option> : masterData.crates.map((c, idx) => <option key={c.id || idx} value={typeof c === 'string' ? c : (c.name || c)}>{typeof c === 'string' ? c : (c.name || c)}</option>)}</select></label>
              <label className="fld"><span>Pallet Type</span><select className="sel" value={form.palletType} onChange={(e) => { const newForm = { ...form, palletType: e.target.value }; if (newForm.lugQuantity && newForm.lugType && newForm.palletType && newForm.grossWeight) { const crate = masterData.crates.find((c) => (typeof c === 'string' ? c : (c.name || c)) === newForm.lugType); const crateWt = crate ? parseFloat(crate.weight || crate.wt || crate.Weight || 0) : 0; const pallet = masterData.pallets.find((p) => (typeof p === 'string' ? p : (p.name || p)) === newForm.palletType); const palletWt = pallet ? parseFloat(pallet.weight || pallet.wt || pallet.Weight || 0) : 0; const lugQty = parseFloat(newForm.lugQuantity) || 0; const grossWt = parseFloat(newForm.grossWeight) || 0; const packagingWeight = (lugQty * crateWt) + palletWt; newForm.netQuantity = (grossWt - packagingWeight >= 0 ? (grossWt - packagingWeight).toFixed(2) : "0.00"); } setForm(newForm); }}><option value="">-- Select pallet type --</option>{masterData.pallets.length === 0 ? <option disabled>No pallet types available</option> : masterData.pallets.map((p, idx) => <option key={p.id || idx} value={typeof p === 'string' ? p : (p.name || p)}>{typeof p === 'string' ? p : (p.name || p)}</option>)}</select></label>
              <label className="fld"><span>Net Quantity (kg)</span><input className="inp" type="number" value={form.netQuantity} readOnly disabled style={{ background: "#f5f5f5", cursor: "not-allowed" }} placeholder="Auto-calculated" /></label>
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, marginTop: 16, justifyContent: "flex-end" }}>
            <button className="btn btn-mini" onClick={() => setForm(null)}>Reset</button>
            <button className="btn btn-primary" onClick={saveRecord}>Save Entry</button>
            {printLabel && <button className="btn btn-success" onClick={() => window.print()}>Print Label</button>}
          </div>
        </Card>
      </>
    );
  }

  return (
    <>
      <div className="page-head">
        <div>
          <h2>Batch Upload</h2>
          <p className="muted">{records.length} total batches recorded</p>
        </div>
        <button className="btn btn-primary" onClick={newRecord}>New Batch</button>
      </div>
      
      {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom: 12 }}>{msg[1]}</p>}
      <div className="grid2">
        <Card title="Period Summary">
          <div className="kpis">
            <Kpi label="Deliveries" value={filtered.length} />
            <Kpi label="Total Quantity" value={kg(totalQty, 1)} sub="kg" />
            <Kpi label="Varieties" value={varieties} />
            <Kpi label="Batches" value={[...new Set(filtered.map((r) => r.batchNumber).filter(Boolean))].length} />
          </div>
        </Card>
        <Card title="Filter">
          <div style={{ display: "flex", gap: 12, flexDirection: "column" }}>
            <label className="fld" style={{ margin: 0 }}>
              <span>From date</span>
              <input className="inp" type="date" value={searchDate} onChange={(e) => setSearchDate(e.target.value)} />
            </label>
          </div>
        </Card>
      </div>
      <Card title="Receiving History" pad={false}>
        {loading ? (
          <p className="muted" style={{ padding: 24 }}>Loading…</p>
        ) : filtered.length === 0 ? (
          <p className="muted" style={{ padding: 24 }}>No batches received yet.</p>
        ) : (
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>Harvest Date</th>
                  <th>Batch Number</th>
                  <th>Section</th>
                  <th>Variety</th>
                  <th>Supplier</th>
                  <th>Net Qty (kg)</th>
                  <th>Lugs</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((r) => (
                  <tr key={r.id}>
                    <td>{dfmt(r.harvestDate)}</td>
                    <td><strong>{r.batchNumber}</strong></td>
                    <td>{r.section}</td>
                    <td>{r.variety}</td>
                    <td>{r.supplier}</td>
                    <td className="r">{kg(r.netQuantity, 1)}</td>
                    <td className="r">{r.lugQuantity}</td>
                    <td style={{ textAlign: "right" }}><button className="btn btn-mini" style={{ background: "#2F7D53" }} onClick={() => setPrintLabel(r)}>Print</button> <button className="btn btn-mini" onClick={() => setForm(r)}>Edit</button> <button className="btn btn-mini btn-danger" onClick={() => deleteRecord(r.id)}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </>
  );
}

/* ================= PRODUCTION PACKING ================= */
/* ================= PACKING OPERATIONS - TABBED INTERFACE ================= */
function ProductionPacking({ me }) {
  const [activeTab, setActiveTab] = useState("work-order");
  const [workOrders, setWorkOrders] = useState([]);
  const [pallets, setPallets] = useState([]);
  const [fruitRecords, setFruitRecords] = useState([]);
  const [masterCustomers, setMasterCustomers] = useState([]);
  const [masterVarieties, setMasterVarieties] = useState([]);
  const [masterPackTypes, setMasterPackTypes] = useState([]);
  const [woForm, setWoForm] = useState(null);
  const [palletForm, setPalletForm] = useState(null);
  const [selectedRecords, setSelectedRecords] = useState(new Set());
  const [selectedWO, setSelectedWO] = useState("");
  const [msg, setMsg] = useState(null);
  const [filterVariety, setFilterVariety] = useState("");
  const [filterStatus, setFilterStatus] = useState("");

  const WO_KEY = "ml-work-orders-v1";
  const PALLET_KEY = "ml-pallets-v1";
  const FR_KEY = "ml-fruit-receiving-v1";
  const CUSTOMER_KEY = "ml-master-customers-v1";
  const VARIETY_KEY = "ml-master-varieties-v1";
  const PACKTYPE_KEY = "ml-master-packtypes-v1";

  // Load data with proper storage handling
  useEffect(() => {
    const loadData = async () => {
      // Load work orders and pallets from localStorage
      const wo = localStorage.getItem(WO_KEY);
      if (wo) {
        try {
          setWorkOrders(JSON.parse(wo));
        } catch (e) {
          console.error("Error parsing work orders:", e);
        }
      }
      
      const pal = localStorage.getItem(PALLET_KEY);
      if (pal) {
        try {
          setPallets(JSON.parse(pal));
        } catch (e) {
          console.error("Error parsing pallets:", e);
        }
      }
      
      const fr = localStorage.getItem(FR_KEY);
      if (fr) {
        try {
          setFruitRecords(JSON.parse(fr));
        } catch (e) {
          console.error("Error parsing fruit records:", e);
        }
      }

      // Load master files from storage
      try {
        // Customers
        const custData = await loadShared(CUSTOMER_KEY);
        if (custData) {
          setMasterCustomers(Array.isArray(custData) ? custData : []);
          console.log("Loaded customers:", custData);
        } else {
          console.warn("No customers found in storage");
        }

        // Varieties
        const varData = await loadShared(VARIETY_KEY);
        if (varData) {
          setMasterVarieties(Array.isArray(varData) ? varData : []);
          console.log("Loaded varieties:", varData);
        } else {
          console.warn("No varieties found in storage");
        }

        // Pack Types
        const packData = await loadShared(PACKTYPE_KEY);
        if (packData) {
          setMasterPackTypes(Array.isArray(packData) ? packData : []);
          console.log("Loaded pack types:", packData);
        } else {
          console.warn("No pack types found in storage");
        }
      } catch (e) {
        console.error("Error loading master files:", e);
      }
    };

    loadData();
  }, []);

  // Save functions
  const saveWO = (data) => {
    localStorage.setItem(WO_KEY, JSON.stringify(data));
    setWorkOrders(data);
  };

  const savePallets = (data) => {
    localStorage.setItem(PALLET_KEY, JSON.stringify(data));
    setPallets(data);
  };

  // Work Order Creation
  const addWorkOrder = () => {
    setWoForm({
      id: "wo-" + Date.now(),
      woNumber: "WO" + new Date().getTime().toString().slice(-6),
      customer: "",
      variety: "",
      packType: "",
      targetQty: "",
      targetKgs: "",
      orderDate: new Date().toISOString().split('T')[0],
      requiredDate: "",
      status: "Pending",
      notes: ""
    });
  };

  const saveWorkOrder = () => {
    if (!woForm.woNumber || !woForm.customer || !woForm.variety) {
      setMsg(["warn", "Fill in WO Number, Customer, and Variety"]);
      return;
    }
    const updated = woForm.id && workOrders.find(w => w.id === woForm.id)
      ? workOrders.map(w => w.id === woForm.id ? woForm : w)
      : [...workOrders, woForm];
    saveWO(updated);
    setWoForm(null);
    setMsg(["ok", "Work Order saved!"]);
  };

  const deleteWO = (id) => {
    saveWO(workOrders.filter(w => w.id !== id));
    setMsg(["ok", "Work Order deleted!"]);
  };

  // Tip to Work Order
  const toggleRecordSelection = (id) => {
    const s = new Set(selectedRecords);
    s.has(id) ? s.delete(id) : s.add(id);
    setSelectedRecords(s);
  };

  const handleTipToWO = () => {
    if (!selectedWO) {
      setMsg(["warn", "Select a work order"]);
      return;
    }
    if (selectedRecords.size === 0) {
      setMsg(["warn", "Select at least one batch"]);
      return;
    }
    setMsg(["ok", `Tipped ${selectedRecords.size} batch(es) to work order!`]);
    setSelectedRecords(new Set());
    setSelectedWO("");
  };

  // Pallet Registration
  const addPallet = () => {
    setPalletForm({
      id: "pal-" + Date.now(),
      palletNumber: "PAL" + new Date().getTime().toString().slice(-6),
      woNumber: "",
      palletType: "",
      quantity: "",
      weight: "",
      registeredDate: new Date().toISOString().split('T')[0],
      status: "Building",
      notes: ""
    });
  };

  const savePallet = () => {
    if (!palletForm.palletNumber || !palletForm.woNumber) {
      setMsg(["warn", "Fill in Pallet Number and Work Order"]);
      return;
    }
    const updated = palletForm.id && pallets.find(p => p.id === palletForm.id)
      ? pallets.map(p => p.id === palletForm.id ? palletForm : p)
      : [...pallets, palletForm];
    savePallets(updated);
    setPalletForm(null);
    setMsg(["ok", "Pallet registered!"]);
  };

  const deletePallet = (id) => {
    savePallets(pallets.filter(p => p.id !== id));
    setMsg(["ok", "Pallet deleted!"]);
  };

  // Filters
  const filteredFruit = fruitRecords.filter(r => {
    if (filterVariety && r.variety !== filterVariety) return false;
    if (filterStatus) {
      const isLinked = selectedRecords.has(r.id);
      if (filterStatus === "linked" && !isLinked) return false;
      if (filterStatus === "unlinked" && isLinked) return false;
    }
    return true;
  });

  const uniqueVarieties = [...new Set(fruitRecords.map(r => r.variety))];
  const filteredPallets = pallets.filter(p => {
    if (filterVariety && !p.woNumber) return false;
    if (filterStatus && p.status !== filterStatus) return false;
    return true;
  });

  const packoutSummary = {
    totalWO: workOrders.length,
    activeWO: workOrders.filter(w => w.status === "In Progress").length,
    totalPallets: pallets.length,
    palletsFull: pallets.filter(p => p.status === "Ready").length,
    totalQtyPacked: pallets.reduce((s, p) => s + parseFloat(p.quantity || 0), 0),
    totalWeightPacked: pallets.reduce((s, p) => s + parseFloat(p.weight || 0), 0),
  };

  // Tab content
  const renderTab = () => {
    if (activeTab === "work-order") {
      if (woForm) {
        return (
          <div style={{ padding: 20 }}>
            <h2>Create Work Order</h2>
            <div className="form2">
              <label className="fld"><span>WO Number *</span><input className="inp" value={woForm.woNumber} onChange={(e) => setWoForm({...woForm, woNumber: e.target.value})} /></label>
              <label className="fld"><span>Customer *</span><select className="inp" value={woForm.customer} onChange={(e) => setWoForm({...woForm, customer: e.target.value})}><option value="">Select customer...</option>{masterCustomers.length === 0 ? <option disabled>No customers in Master Files</option> : masterCustomers.map((c, idx) => <option key={c.id || idx} value={typeof c === 'string' ? c : (c.name || c.code || c)}>{typeof c === 'string' ? c : (c.name || c.code || c)}</option>)}</select></label>
              <label className="fld"><span>Variety *</span><select className="inp" value={woForm.variety} onChange={(e) => setWoForm({...woForm, variety: e.target.value})}><option value="">Select variety...</option>{masterVarieties.length === 0 ? <option disabled>No varieties in Master Files</option> : masterVarieties.map((v, idx) => <option key={v.id || idx} value={typeof v === 'string' ? v : (v.name || v.code || v)}>{typeof v === 'string' ? v : (v.name || v.code || v)}</option>)}</select></label>
              <label className="fld"><span>Pack Type</span><select className="inp" value={woForm.packType} onChange={(e) => setWoForm({...woForm, packType: e.target.value})}><option value="">Select type...</option>{masterPackTypes.length === 0 ? <option disabled>No pack types in Master Files</option> : masterPackTypes.map((p, idx) => <option key={p.id || idx} value={typeof p === 'string' ? p : (p.name || p.code || p)}>{typeof p === 'string' ? p : (p.name || p.code || p)}</option>)}</select></label>
              <label className="fld"><span>Target Quantity</span><input className="inp" type="number" value={woForm.targetQty} onChange={(e) => setWoForm({...woForm, targetQty: e.target.value})} placeholder="Units" /></label>
              <label className="fld"><span>Target Weight (kg)</span><input className="inp" type="number" step="0.01" value={woForm.targetKgs} onChange={(e) => setWoForm({...woForm, targetKgs: e.target.value})} placeholder="0.00" /></label>
              <label className="fld"><span>Required Date</span><input className="inp" type="date" value={woForm.requiredDate} onChange={(e) => setWoForm({...woForm, requiredDate: e.target.value})} /></label>
              <label className="fld"><span>Status</span><select className="inp" value={woForm.status} onChange={(e) => setWoForm({...woForm, status: e.target.value})}><option value="Pending">Pending</option><option value="In Progress">In Progress</option><option value="Completed">Completed</option></select></label>
              <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Notes</span><textarea className="inp" value={woForm.notes} onChange={(e) => setWoForm({...woForm, notes: e.target.value})} style={{ minHeight: 60 }} /></label>
            </div>
            {msg && <p className={"savedmsg " + msg[0]}>{msg[1]}</p>}
            <button className="btn btn-primary" onClick={saveWorkOrder}>Save Work Order</button>
            <button className="btn" onClick={() => setWoForm(null)}>Cancel</button>
          </div>
        );
      }
      return (
        <>
          <div style={{ marginBottom: 16 }}>
            <button className="btn btn-primary" onClick={addWorkOrder}>+ New Work Order</button>
          </div>
          {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom: 12 }}>{msg[1]}</p>}
          <Card title={`Work Orders (${workOrders.length})`}>
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr><th>WO Number</th><th>Customer</th><th>Variety</th><th>Pack Type</th><th className="r">Target Qty</th><th className="r">Target (kg)</th><th>Required</th><th>Status</th><th>Actions</th></tr>
                </thead>
                <tbody>
                  {workOrders.map(w => (
                    <tr key={w.id}>
                      <td><strong>{w.woNumber}</strong></td>
                      <td>{w.customer}</td>
                      <td>{w.variety}</td>
                      <td>{w.packType}</td>
                      <td className="r">{w.targetQty}</td>
                      <td className="r">{parseFloat(w.targetKgs || 0).toFixed(2)}</td>
                      <td>{w.requiredDate}</td>
                      <td><span style={{ padding: "2px 6px", borderRadius: 3, background: w.status === "Completed" ? "#e8f5e9" : w.status === "In Progress" ? "#fff3e0" : "#f5f5f5", fontSize: 11 }}>{w.status}</span></td>
                      <td style={{ textAlign: "right" }}><button className="btn btn-mini" onClick={() => setWoForm(w)}>Edit</button> <button className="btn btn-mini btn-danger" onClick={() => deleteWO(w.id)}>Del</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      );
    }

    if (activeTab === "tip-fruit") {
      return (
        <>
          <h3>Tip In-Stock Fruit to Work Order</h3>
          <Card title="Selection">
            <div className="form2">
              <label className="fld"><span>Select Work Order *</span>
                <select className="inp" value={selectedWO} onChange={(e) => setSelectedWO(e.target.value)}>
                  <option value="">Select work order...</option>
                  {workOrders.filter(w => w.status !== "Completed").map(w => <option key={w.id} value={w.id}>{w.woNumber} - {w.customer}</option>)}
                </select>
              </label>
              <label className="fld"><span>Filter by Variety</span>
                <select className="inp" value={filterVariety} onChange={(e) => setFilterVariety(e.target.value)}>
                  <option value="">All varieties</option>
                  {uniqueVarieties.map(v => <option key={v} value={v}>{v}</option>)}
                </select>
              </label>
            </div>
            {selectedRecords.size > 0 && (
              <div style={{ padding: 12, background: "#e3f2fd", borderRadius: 4, marginBottom: 12 }}>
                <strong>{selectedRecords.size}</strong> batch(es) selected
                <button className="btn btn-primary" style={{ marginLeft: 12, float: "right" }} onClick={handleTipToWO} disabled={!selectedWO}>Tip to Work Order</button>
              </div>
            )}
          </Card>
          {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom: 12 }}>{msg[1]}</p>}
          <Card title={`Available Fruit Batches (${filteredFruit.length})`}>
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr><th style={{ width: 30 }}><input type="checkbox" checked={selectedRecords.size === filteredFruit.length && filteredFruit.length > 0} onChange={(e) => setSelectedRecords(e.target.checked ? new Set(filteredFruit.map(r => r.id)) : new Set())} /></th><th>Batch #</th><th>Harvest Date</th><th>Supplier</th><th>Variety</th><th>Block</th><th className="r">Qty (kg)</th></tr>
                </thead>
                <tbody>
                  {filteredFruit.map(r => (
                    <tr key={r.id} style={{ background: selectedRecords.has(r.id) ? "#e3f2fd" : "" }}>
                      <td><input type="checkbox" checked={selectedRecords.has(r.id)} onChange={() => toggleRecordSelection(r.id)} /></td>
                      <td><strong>{r.batchNumber}</strong></td>
                      <td>{r.harvestDate}</td>
                      <td>{r.supplier}</td>
                      <td>{r.variety}</td>
                      <td>{r.block}</td>
                      <td className="r">{parseFloat(r.netQuantity || 0).toFixed(2)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      );
    }

    if (activeTab === "pallet-buildup") {
      if (palletForm) {
        return (
          <div style={{ padding: 20 }}>
            <h2>Register Pallet</h2>
            <div className="form2">
              <label className="fld"><span>Pallet Number *</span><input className="inp" value={palletForm.palletNumber} onChange={(e) => setPalletForm({...palletForm, palletNumber: e.target.value})} /></label>
              <label className="fld"><span>Work Order *</span>
                <select className="inp" value={palletForm.woNumber} onChange={(e) => setPalletForm({...palletForm, woNumber: e.target.value})}>
                  <option value="">Select work order...</option>
                  {workOrders.map(w => <option key={w.id} value={w.woNumber}>{w.woNumber} - {w.customer}</option>)}
                </select>
              </label>
              <label className="fld"><span>Pallet Type</span><input className="inp" value={palletForm.palletType} onChange={(e) => setPalletForm({...palletForm, palletType: e.target.value})} placeholder="E.g., Air Pallet, Wood Pallet" /></label>
              <label className="fld"><span>Quantity (units)</span><input className="inp" type="number" value={palletForm.quantity} onChange={(e) => setPalletForm({...palletForm, quantity: e.target.value})} /></label>
              <label className="fld"><span>Weight (kg)</span><input className="inp" type="number" step="0.01" value={palletForm.weight} onChange={(e) => setPalletForm({...palletForm, weight: e.target.value})} /></label>
              <label className="fld"><span>Status</span><select className="inp" value={palletForm.status} onChange={(e) => setPalletForm({...palletForm, status: e.target.value})}><option value="Building">Building</option><option value="Full">Full</option><option value="Ready">Ready for Dispatch</option></select></label>
              <label className="fld" style={{ gridColumn: "1 / -1" }}><span>Notes</span><textarea className="inp" value={palletForm.notes} onChange={(e) => setPalletForm({...palletForm, notes: e.target.value})} style={{ minHeight: 60 }} /></label>
            </div>
            {msg && <p className={"savedmsg " + msg[0]}>{msg[1]}</p>}
            <button className="btn btn-primary" onClick={savePallet}>Save Pallet</button>
            <button className="btn" onClick={() => setPalletForm(null)}>Cancel</button>
          </div>
        );
      }
      return (
        <>
          <div style={{ marginBottom: 16 }}>
            <button className="btn btn-primary" onClick={addPallet}>+ New Pallet</button>
          </div>
          {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom: 12 }}>{msg[1]}</p>}
          <Card title={`Pallets (${pallets.length})`}>
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr><th>Pallet #</th><th>Work Order</th><th>Type</th><th className="r">Qty</th><th className="r">Weight (kg)</th><th>Date</th><th>Status</th><th>Actions</th></tr>
                </thead>
                <tbody>
                  {pallets.map(p => (
                    <tr key={p.id}>
                      <td><strong>{p.palletNumber}</strong></td>
                      <td>{p.woNumber}</td>
                      <td>{p.palletType}</td>
                      <td className="r">{p.quantity}</td>
                      <td className="r">{parseFloat(p.weight || 0).toFixed(2)}</td>
                      <td>{p.registeredDate}</td>
                      <td><span style={{ padding: "2px 6px", borderRadius: 3, background: p.status === "Ready" ? "#e8f5e9" : p.status === "Full" ? "#fff3e0" : "#f5f5f5", fontSize: 11 }}>{p.status}</span></td>
                      <td style={{ textAlign: "right" }}><button className="btn btn-mini" onClick={() => setPalletForm(p)}>Edit</button> <button className="btn btn-mini btn-danger" onClick={() => deletePallet(p.id)}>Del</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      );
    }

    if (activeTab === "pallet-report") {
      return (
        <>
          <h3>Registered Pallets Report</h3>
          <Card title="Filters">
            <div className="form2">
              <label className="fld"><span>Filter by Status</span>
                <select className="inp" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                  <option value="">All statuses</option>
                  <option value="Building">Building</option>
                  <option value="Full">Full</option>
                  <option value="Ready">Ready for Dispatch</option>
                </select>
              </label>
            </div>
          </Card>
          <Card title={`Pallet Report (${filteredPallets.length})`}>
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr><th>Pallet #</th><th>Work Order</th><th>Type</th><th className="r">Qty</th><th className="r">Weight (kg)</th><th>Registered Date</th><th>Status</th><th>Notes</th></tr>
                </thead>
                <tbody>
                  {filteredPallets.map(p => (
                    <tr key={p.id}>
                      <td><strong>{p.palletNumber}</strong></td>
                      <td>{p.woNumber}</td>
                      <td>{p.palletType}</td>
                      <td className="r">{p.quantity}</td>
                      <td className="r">{parseFloat(p.weight || 0).toFixed(2)}</td>
                      <td>{p.registeredDate}</td>
                      <td><span style={{ padding: "2px 6px", borderRadius: 3, background: p.status === "Ready" ? "#e8f5e9" : p.status === "Full" ? "#fff3e0" : "#f5f5f5", fontSize: 11 }}>{p.status}</span></td>
                      <td style={{ fontSize: 12, color: "#666" }}>{p.notes}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      );
    }

    if (activeTab === "packout-summary") {
      return (
        <>
          <h3>Packout Summary</h3>
          <Card title="Operations Summary">
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 16 }}>
              <div style={{ padding: 16, background: "#e3f2fd", borderRadius: 6 }}>
                <div style={{ fontSize: 12, color: "#1976d2", marginBottom: 6 }}>Total Work Orders</div>
                <div style={{ fontSize: 28, fontWeight: "bold", color: "#1976d2" }}>{packoutSummary.totalWO}</div>
              </div>
              <div style={{ padding: 16, background: "#fff3e0", borderRadius: 6 }}>
                <div style={{ fontSize: 12, color: "#f57c00", marginBottom: 6 }}>Active Work Orders</div>
                <div style={{ fontSize: 28, fontWeight: "bold", color: "#f57c00" }}>{packoutSummary.activeWO}</div>
              </div>
              <div style={{ padding: 16, background: "#e8f5e9", borderRadius: 6 }}>
                <div style={{ fontSize: 12, color: "#388e3c", marginBottom: 6 }}>Total Pallets</div>
                <div style={{ fontSize: 28, fontWeight: "bold", color: "#388e3c" }}>{packoutSummary.totalPallets}</div>
              </div>
              <div style={{ padding: 16, background: "#f3e5f5", borderRadius: 6 }}>
                <div style={{ fontSize: 12, color: "#7b1fa2", marginBottom: 6 }}>Ready Pallets</div>
                <div style={{ fontSize: 28, fontWeight: "bold", color: "#7b1fa2" }}>{packoutSummary.palletsFull}</div>
              </div>
              <div style={{ padding: 16, background: "#fce4ec", borderRadius: 6 }}>
                <div style={{ fontSize: 12, color: "#c2185b", marginBottom: 6 }}>Total Qty Packed</div>
                <div style={{ fontSize: 28, fontWeight: "bold", color: "#c2185b" }}>{packoutSummary.totalQtyPacked}</div>
              </div>
              <div style={{ padding: 16, background: "#ffebee", borderRadius: 6 }}>
                <div style={{ fontSize: 12, color: "#d32f2f", marginBottom: 6 }}>Total Weight (kg)</div>
                <div style={{ fontSize: 28, fontWeight: "bold", color: "#d32f2f" }}>{packoutSummary.totalWeightPacked.toFixed(1)}</div>
              </div>
            </div>
          </Card>
          <Card title="Work Orders Summary">
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr><th>WO #</th><th>Customer</th><th>Variety</th><th className="r">Target Qty</th><th className="r">Target (kg)</th><th>Due</th><th>Status</th></tr>
                </thead>
                <tbody>
                  {workOrders.map(w => (
                    <tr key={w.id}>
                      <td><strong>{w.woNumber}</strong></td>
                      <td>{w.customer}</td>
                      <td>{w.variety}</td>
                      <td className="r">{w.targetQty}</td>
                      <td className="r">{parseFloat(w.targetKgs || 0).toFixed(2)}</td>
                      <td>{w.requiredDate}</td>
                      <td><span style={{ padding: "2px 6px", borderRadius: 3, background: w.status === "Completed" ? "#e8f5e9" : w.status === "In Progress" ? "#fff3e0" : "#f5f5f5", fontSize: 11 }}>{w.status}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      );
    }
  };

  const tabs = [
    { id: "work-order", label: "Work Order Creation" },
    { id: "tip-fruit", label: "Tip to Work Order" },
    { id: "pallet-buildup", label: "Pallet Registration & Build up" },
    { id: "pallet-report", label: "Registered Pallets Report" },
    { id: "packout-summary", label: "Packout Summary" },
  ];

  return (
    <div>
      <h2>Packing Operations</h2>
      <p className="muted">Manage work orders, pallets, and packing operations</p>
      
      <div style={{ display: "flex", gap: 0, borderBottom: "1px solid #ddd", marginBottom: 20, overflowX: "auto" }}>
        {tabs.map(tab => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id)}
            style={{
              padding: "12px 16px",
              border: "none",
              background: activeTab === tab.id ? "#2F7D53" : "transparent",
              color: activeTab === tab.id ? "white" : "#666",
              cursor: "pointer",
              fontSize: 13,
              fontWeight: activeTab === tab.id ? "600" : "500",
              borderBottom: activeTab === tab.id ? "3px solid #2F7D53" : "none",
              whiteSpace: "nowrap"
            }}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {renderTab()}
    </div>
  );
}

/* ================= COMPLIANCE ================= */
class ComplianceErrorBoundary extends React.Component {
  constructor(p) { super(p); this.state = { err: null, info: null }; }
  static getDerivedStateFromError(e) { return { err: e }; }
  componentDidCatch(e, info) { this.setState({ info: info?.componentStack?.slice(0, 300) }); }
  render() {
    if (this.state.err) return (
      <div style={{ padding: 24 }}>
        <h3 style={{ color: "var(--red)" }}>Something went wrong in Compliance</h3>
        <p className="muted" style={{ marginBottom: 6 }}><b>Error:</b> {String(this.state.err?.message || this.state.err)}</p>
        {this.state.info && <pre style={{ fontSize: 10, background: "#F6F6F9", padding: 8, borderRadius: 6, overflow: "auto", maxHeight: 120, marginBottom: 10 }}>{this.state.info}</pre>}
        <button className="btn btn-mini" onClick={() => this.setState({ err: null, info: null })}>Try again</button>
        <p className="muted small" style={{ marginTop: 8 }}>Screenshot this (including the grey text above) and send it to Claude.</p>
      </div>
    );
    return this.props.children;
  }
}

function Compliance({ me }) {
  const [records, setRecords] = useState([]);  // always array
  const [loading, setLoading] = useState(true);
  const [active, setActive] = useState(null); // { form, readOnly }
  const [selType, setSelType] = useState("elifab");
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    loadShared(COMPLIANCE_KEY)
      .then((d) => {
        console.log("[Compliance] raw from storage:", typeof d, JSON.stringify(d)?.slice(0, 200));
        let safe = [];
        if (Array.isArray(d)) safe = d.filter((x) => x && typeof x === "object" && x.id);
        else if (d && typeof d === "object" && !Array.isArray(d)) safe = Object.values(d).filter((x) => x && x.id);
        console.log("[Compliance] safe records:", safe.length);
        setRecords(safe);
        setLoading(false);
      })
      .catch((e) => {
        console.error("[Compliance] load error:", e);
        setRecords([]);
        setLoading(false);
      });
  }, []);

  const canFill = me.role === "admin" || me.role === "qc" || me.role === "entry";
  const canApprove = me.role === "admin" || me.role === "qc";
  const editForm = (r) => setActive({ form: r, readOnly: false, editing: true });
  const FORM_TYPES = {
    elifab:       { label: "Elifab Machine Cleaning & Inspection", new: () => newElifabForm(me),              pdfName: (f) => "Elifab_Cleaning_" + f.date },
    coldroom:     { label: "Coldroom Temperature Check",            new: () => newCrForm(me),                  pdfName: (f) => "Coldroom_Temp_Wk" + f.week },
    crate:        { label: "Crate Cleaning Record",                 new: () => newCrateForm(me),               pdfName: (f) => "Crate_Cleaning_" + f.month },
    label:        { label: "Labelling Roll Approval",               new: () => newLabelForm(me),               pdfName: (f) => "Labelling_Roll_" + f.date },
    pallet:       { label: "Pallet Composition",                    new: () => newPalletForm(me),              pdfName: (f) => "Pallet_Composition_" + f.date },
    hygieneField: { label: "Personal Hygiene - Field",              new: () => makeHygieneForm("hygieneField", me), pdfName: (f) => "Hygiene_Field_" + f.date },
    hygienePack:  { label: "Personal Hygiene - Packhouse",          new: () => makeHygieneForm("hygienePack", me),  pdfName: (f) => "Hygiene_Packhouse_" + f.date },
    qcReport:     { label: "Post-Machine QC Report",                new: () => newQcForm(me),                  pdfName: (f) => "QC_Report_" + f.date },
    scale:        { label: "Scale Calibration Check",               new: () => newScaleForm(me),               pdfName: (f) => "Scale_Calibration_" + f.date },
    truck:        { label: "Truck Hygiene and Temp Check",          new: () => newTruckForm(me),               pdfName: (f) => "Truck_Hygiene_" + f.month },
    visitor:      { label: "Visitor Questionnaire",                 new: () => newVisitorForm(me),             pdfName: (f) => "Visitor_" + f.date },
  };
  const newForm = (type) => setActive({ form: FORM_TYPES[type] ? FORM_TYPES[type].new() : newElifabForm(me), readOnly: false });
  const viewForm = (r) => setActive({ form: r, readOnly: true });

  const compressForm = (form) => {
    // re-compress any photo fields to JPEG 0.7 to keep storage lean
    const PHOTO_KEYS = ["punnetFirstPhoto","punnetLastPhoto","boxFirstPhoto","boxLastPhoto"];
    let out = { ...form };
    PHOTO_KEYS.forEach((k) => {
      if (!out[k] || !out[k].startsWith("data:image")) return;
      try {
        const img = new Image();
        img.src = out[k];
        const canvas = document.createElement("canvas");
        const MAX = 800;
        const scale = img.width > MAX ? MAX / img.width : 1;
        canvas.width = (img.width || 800) * scale;
        canvas.height = (img.height || 600) * scale;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        const compressed = canvas.toDataURL("image/jpeg", 0.7);
        if (compressed.length < out[k].length) out[k] = compressed;
      } catch (e) { /* leave as-is on error */ }
    });
    return out;
  };

  const save = async (form) => {
    setBusy(true);
    const isEdit = !!(active && active.editing);
    const cf = compressForm(form);
    const saved = isEdit
      ? { ...cf, editedBy: me.name, editedUsername: me.username, editedAt: new Date().toISOString(), approved: null, approvedBy: null, approvedAt: null }
      : { ...cf, submitted: new Date().toISOString(), submittedBy: me.name, submittedUsername: me.username };
    const cur = Array.isArray(records) ? records : [];
    const next = isEdit
      ? cur.map((r) => r.id === saved.id ? saved : r)
      : [...cur, saved];
    const ok = await saveShared(COMPLIANCE_KEY, next);
    setRecords(Array.isArray(next) ? next : []);
    setActive(null);
    setMsg(ok ? ["ok", isEdit ? "Form updated." : "Form submitted and saved."] : ["warn", "Saved locally but shared storage is unavailable."]);
    setBusy(false);
  };

  const deleteComplianceForm = async (r) => {
    if (!window.confirm("Delete this form permanently? This cannot be undone.")) return;
    const next = (Array.isArray(records) ? records : []).filter((x) => x.id !== r.id);
    const ok = await saveShared(COMPLIANCE_KEY, next);
    setRecords(Array.isArray(next) ? next : []);
    if (active && active.form.id === r.id) setActive(null);
    setMsg(ok ? ["ok", "Form deleted."] : ["warn", "Deleted for this session, but shared storage is unavailable."]);
  };

    const approveForm = async (r) => {
    const approved = { ...r, approved: new Date().toISOString(), approvedBy: me.name, approvedUsername: me.username };
    const next = (Array.isArray(records) ? records : []).map((x) => x.id === r.id ? approved : x);
    const ok = await saveShared(COMPLIANCE_KEY, next);
    setRecords(Array.isArray(next) ? next : []);
    if (active && active.form.id === r.id) setActive({ form: approved, readOnly: true });
    setMsg(ok ? ["ok", "Form approved by " + me.name + "."] : ["warn", "Saved locally but shared storage is unavailable."]);
  };

  const downloadPdf = async (r) => {
    try {
      await loadSnapLibs();
      const el = document.getElementById("comp-print-" + r.id);
      if (!el) return;
      const canvas = await window.html2canvas(el, { scale: 2, backgroundColor: "#ffffff", useCORS: true, width: el.scrollWidth, windowWidth: el.scrollWidth });
      const { jsPDF } = window.jspdf;
      const landscape = canvas.width > canvas.height;
      const pdf = new jsPDF({ orientation: landscape ? "landscape" : "portrait", unit: "mm", format: "a4" });
      const M = 8, availW = (landscape ? 297 : 210) - 2 * M, availH = (landscape ? 210 : 297) - 2 * M;
      const ratio = canvas.height / canvas.width;
      let w = availW, h = availW * ratio;
      if (h > availH) { h = availH; w = availH / ratio; }
      pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", M + (availW - w) / 2, M, w, h);
      pdf.save("Merrylight_" + (r.type === 'coldroom' ? 'Coldroom_Temp_Wk' + r.week : 'Elifab_Cleaning_' + r.date) + ".pdf");
    } catch (e) { alert("PDF failed: " + (e && e.message)); }
  };

  if (active) {
    if (!active.form || !active.form.type) {
      return (
        <div style={{ padding: 24 }}>
          <p className="muted">This form record is missing data. It may have been corrupted during save.</p>
          <button className="btn btn-mini" onClick={() => setActive(null)}>Back to list</button>
        </div>
      );
    }
    return (
      <>
        <div className="page-head">
          <div>
            <h2>{(FORM_TYPES[active.form.type] || {}).label || active.form.type}</h2>
            <p className="muted">
              {active.readOnly
                ? (active.form.submitted ? "Submitted " + dfmt(active.form.submitted.slice(0,10)) + " by " + active.form.submittedBy : "Draft")
                : (active.editing ? "Editing" : "New form")}
              {active.form.approved
                ? <span className="badge ok" style={{ marginLeft: 8 }}>Approved by {active.form.approvedBy}</span>
                : active.form.submitted
                  ? <span className="badge" style={{ marginLeft: 8, background:"#FFF3CD",color:"#856404" }}>Pending approval</span>
                  : null}
            </p>
          </div>
          <div style={{ display:"flex", gap:6 }}>
            {active.readOnly && !active.form.approved && canFill && (
              <button className="btn btn-mini" onClick={() => editForm(active.form)}>Edit</button>
            )}
            {active.readOnly && !active.form.approved && canApprove && (
              <button className="btn btn-primary" style={{ width:"auto", padding:"7px 12px" }} onClick={() => approveForm(active.form)}>Approve</button>
            )}
            {active.form.approved && active.readOnly && (
              <button className="btn btn-mini" onClick={() => downloadPdf(active.form)}>Download PDF</button>
            )}
            {me.role === "admin" && (
              <button className="btn btn-danger" style={{ padding:"7px 12px", width:"auto" }} onClick={() => deleteComplianceForm(active.form)}>Delete form</button>
            )}
            <button className="btn btn-mini" onClick={() => setActive(null)}>Back to list</button>
          </div>
        </div>
        <Card>
          {(() => {
            const _p = { form: active.form, onChange: (f) => setActive({ ...active, form: f }), onSave: save, onCancel: () => setActive(null), me, readOnly: active.readOnly };
            const _t = active.form.type;
            if (_t === "coldroom") return <ColdRoomForm {..._p} />;
            if (_t === "crate") return <CrateForm {..._p} />;
            if (_t === "label") return <LabelForm {..._p} />;
            if (_t === "pallet") return <PalletForm {..._p} />;
            if (_t === "hygieneField" || _t === "hygienePack") return <HygieneForm {..._p} />;
            if (_t === "qcReport") return <QcForm {..._p} />;
            if (_t === "scale") return <ScaleForm {..._p} />;
            if (_t === "truck") return <TruckForm {..._p} />;
            if (_t === "visitor") return <VisitorForm {..._p} />;
            return <ElifabForm {..._p} />;
          })()}
        </Card>
        {active.readOnly && (
          <div style={{ marginTop: 8 }}>
            <button className="btn btn-mini" onClick={() => downloadPdf(active.form)}>Download PDF</button>
            <div id={"comp-print-" + active.form.id} style={{ position: "absolute", left: -9999, top: 0, width: 900, background: "#fff", padding: 24 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, borderBottom: "2px solid #4C46A8", paddingBottom: 8, marginBottom: 14 }}>
                <img src={LOGO} alt="" style={{ width: 36, height: 36, borderRadius: "50%" }} />
                <div>
                  <div style={{ fontWeight: 800, fontSize: 15, fontFamily: "sans-serif" }}>MERRYLIGHT PACKHOUSE - {((FORM_TYPES[active.form.type] || {}).label || active.form.type).toUpperCase()}</div>
                  <div style={{ fontSize: 11, color: "#6E6D85" }}>{active.form.type === "coldroom" ? "Week: " + active.form.week : "Date: " + active.form.date + " · Completed by: " + active.form.completedBy}</div>
                  <div style={{ fontSize: 11, color: "#6E6D85" }}>Submitted: {active.form.submitted?.slice(0,16).replace("T"," ")} · Submitted by: <b>{active.form.submittedBy}</b> ({active.form.submittedUsername || ""})</div>
                  {active.form.approved && <div style={{ fontSize: 11, color:"#1A5C2A", fontWeight:700 }}>✓ Approved by {active.form.approvedBy} ({active.form.approvedUsername || ""}) · {active.form.approved?.slice(0,16).replace("T"," ")}</div>}
                  {active.form.signature && <div style={{ marginTop: 6 }}><div style={{ fontSize: 10, color: "#6E6D85", marginBottom: 2 }}>Signature:</div><img src={active.form.signature} style={{ height: 50, border: "1px solid #D8D9E6", borderRadius: 6 }} alt="Signature" /></div>}
                </div>
              </div>
          {(() => {
            const _p = { form: active.form, onChange: (f) => () => {}, onSave: () => {}, onCancel: () => () => {}, me, readOnly: true };
            const _t = active.form.type;
            if (_t === "coldroom") return <ColdRoomForm {..._p} />;
            if (_t === "crate") return <CrateForm {..._p} />;
            if (_t === "label") return <LabelForm {..._p} />;
            if (_t === "pallet") return <PalletForm {..._p} />;
            if (_t === "hygieneField" || _t === "hygienePack") return <HygieneForm {..._p} />;
            if (_t === "qcReport") return <QcForm {..._p} />;
            if (_t === "scale") return <ScaleForm {..._p} />;
            if (_t === "truck") return <TruckForm {..._p} />;
            if (_t === "visitor") return <VisitorForm {..._p} />;
            return <ElifabForm {..._p} />;
          })()}
            </div>
          </div>
        )}
      </>
    );
  }

  return (
    <>
      <div className="page-head">
        <div><h2>{(FORM_TYPES[selType] || {}).label || "Compliance"}</h2><p className="muted">{records.filter((r) => r && r.type === selType).length} submission{records.filter((r) => r && r.type === selType).length === 1 ? "" : "s"}</p></div>
        {canFill && (
        <div style={{ display:"flex", gap:8 }}>
          <select className="inp" value={selType} style={{ width: 260 }} onChange={(e) => setSelType(e.target.value)}>
            {Object.entries(FORM_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
          </select>
          <button className="btn btn-primary" style={{ width:"auto" }} onClick={() => newForm(selType)}>New form</button>
        </div>
      )}
      </div>
      {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom: 12 }}>{msg[1]}</p>}
      <Card pad={false}>
        <div className="tbl-wrap">
          <table>
            <thead><tr>
<th>Type</th><th>Date / Week</th><th>Completed by</th><th>Verified by</th>
              <th className="r">Pass</th><th className="r">Fail</th><th className="r">N/A</th>
              <th>Status</th><th>Submitted by</th><th className="r">Actions</th>
            </tr></thead>
            <tbody>
              {loading && <tr><td colSpan={9} className="muted center">Loading…</td></tr>}
              {!loading && !records.filter((r) => r && r.type === selType).length && <tr><td colSpan={9} className="muted center">No {(FORM_TYPES[selType] || {}).label || "form"} submissions yet.</td></tr>}
              {!loading && records.filter((r) => r && r.id && r.type === selType).sort((a,b) => (b.date||b.submitted||"").localeCompare(a.date||a.submitted||"")).map((r) => (
                <tr key={r.id}>
                  <td><span className="badge">{(FORM_TYPES[r.type] || {}).label?.split(" ")[0] || r.type}</span></td>
                  <td>{r.type === "coldroom" ? "Wk " + r.week : dfmt(r.date)}</td>
                  <td className="strong">{r.type === "coldroom" ? (r.readings && r.readings[0] && r.readings[0].doneBy) || "—" : r.completedBy}</td>
                  <td>{r.type === "coldroom" ? "—" : (r.verifiedBy || "—")}</td>
                  <td className="r">{r.type === "coldroom" ? (r.readings ? r.readings.length + " day(s)" : "0 day(s)")
                    : r.type === "truck" ? (r.rows ? r.rows.length + " row(s)" : "—")
                    : r.type === "crate" ? (r.rows ? r.rows.length + " row(s)" : "—")
                    : r.type === "pallet" ? (r.lines ? r.lines.length + " line(s)" : "—")
                    : r.type === "scale" || r.type === "label" || r.type === "visitor" ? "—"
                    : (Array.isArray(r.items) ? r.items.filter(x => x.status === "Pass").length : "—")}</td>
                  <td className={"r " + (Array.isArray(r.items) && r.items.filter(x => x.status === "Fail").length ? "neg strong" : "")}>
                    {Array.isArray(r.items) ? r.items.filter(x => x.status === "Fail").length : "—"}</td>
                  <td className="r muted">{r.type === "coldroom" ? (r.corrective ? r.corrective.length + " OOS" : "0 OOS")
                    : Array.isArray(r.items) ? r.items.filter(x => x.status === "N/A").length : "—"}</td>
                  <td>{r.approved
                    ? <span className="badge ok">Approved</span>
                    : <span className="badge" style={{ background:"#FFF3CD",color:"#856404" }}>Pending</span>}
                  </td>
                  <td className="muted small">{r.submittedBy}</td>
                  <td className="r">
                    <button className="btn btn-mini" onClick={() => viewForm(r)}>View</button>{" "}
                    {!r.approved && canFill && <><button className="btn btn-mini" onClick={() => editForm(r)}>Edit</button>{" "}</>}
                    {!r.approved && canApprove && (
                      <button className="btn btn-mini" style={{ background:"var(--bloom)",color:"#fff",border:0 }} onClick={() => approveForm(r)}>Approve</button>
                    )}{" "}
                    {r.approved && <button className="btn btn-mini" onClick={() => downloadPdf(r)}>PDF</button>}
                    {me.role === "admin" && <>{" "}<button className="btn btn-mini btn-danger" style={{ padding:"3px 8px" }} onClick={() => deleteComplianceForm(r)}>Delete</button></>}
                    <div id={"comp-print-" + r.id} style={{ position: "absolute", left: -9999, top: 0, width: 900, background: "#fff", padding: 24 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 12, borderBottom: "2px solid #4C46A8", paddingBottom: 8, marginBottom: 14 }}>
                        <img src={LOGO} alt="" style={{ width: 36, height: 36, borderRadius: "50%" }} />
                        <div>
                          <div style={{ fontWeight: 800, fontSize: 15, fontFamily: "sans-serif" }}>MERRYLIGHT PACKHOUSE — {r.type === "coldroom" ? "COLDROOM TEMPERATURE CHECK" : "ELIFAB MACHINE CLEANING & INSPECTION"}</div>
                          <div style={{ fontSize: 11, color: "#6E6D85" }}>{r.type === "coldroom" ? "Week: " + r.week : "Date: " + r.date + " · Week: " + r.week + " · Completed by: " + r.completedBy + " · Verified by: " + (r.verifiedBy || "—")}</div>
                          <div style={{ fontSize: 11, color: "#6E6D85" }}>Submitted: {r.submitted?.slice(0,16).replace("T"," ")} · Submitted by: <b>{r.submittedBy}</b> ({r.submittedUsername || ""})</div>
                          {r.approved && <div style={{ fontSize: 11, color:"#1A5C2A", fontWeight:700 }}>✓ Approved by {r.approvedBy} ({r.approvedUsername || ""}) · {r.approved?.slice(0,16).replace("T"," ")}</div>}
                          {r.signature && <div style={{ marginTop: 4 }}><div style={{ fontSize: 10, color: "#6E6D85", marginBottom: 2 }}>Signature:</div><img src={r.signature} style={{ height: 50, border: "1px solid #D8D9E6", borderRadius: 6 }} alt="Signature" /></div>}
                        </div>
                      </div>
                      {r.type === "coldroom"
                        ? <ColdRoomForm form={r} onChange={() => {}} onSave={() => {}} onCancel={() => {}} me={me} readOnly={true} />
                        : <ElifabForm form={r} onChange={() => {}} onSave={() => {}} onCancel={() => {}} me={me} readOnly={true} />}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </>
  );
}

/* ================= HUMAN RESOURCES ================= */
const HR_KEY = "ml-hr-v1";
const HR_FORM_TYPES = {}; // populated as forms are added

class HRErrorBoundary extends React.Component {
  constructor(p) { super(p); this.state = { err: null, info: null }; }
  static getDerivedStateFromError(e) { return { err: e }; }
  componentDidCatch(e, info) { this.setState({ info: info?.componentStack?.slice(0, 300) }); }
  render() {
    if (this.state.err) return (
      <div style={{ padding: 24 }}>
        <h3 style={{ color: "var(--red)" }}>Something went wrong in Human Resources</h3>
        <p className="muted" style={{ marginBottom: 6 }}><b>Error:</b> {String(this.state.err?.message || this.state.err)}</p>
        {this.state.info && <pre style={{ fontSize: 10, background: "#F6F6F9", padding: 8, borderRadius: 6, overflow: "auto", maxHeight: 120, marginBottom: 10 }}>{this.state.info}</pre>}
        <button className="btn btn-mini" onClick={() => this.setState({ err: null, info: null })}>Try again</button>
        <p className="muted small" style={{ marginTop: 8 }}>Screenshot this and send it to Claude.</p>
      </div>
    );
    return this.props.children;
  }
}

function HumanResources({ me }) {
  const [records, setRecords] = useState([]);
  const [loading, setLoading] = useState(true);
  const [active, setActive] = useState(null);
  const [selType, setSelType] = useState(null);
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    loadShared(HR_KEY)
      .then((d) => { setRecords(Array.isArray(d) ? d.filter((x) => x && x.id) : []); setLoading(false); })
      .catch(() => { setRecords([]); setLoading(false); });
  }, []);

  const canFill    = me.role === "admin" || me.role === "hrOfficer" || me.role === "hrManager";
  const canApprove = me.role === "admin" || me.role === "hrManager";
  const canEditContract = me.role === "admin" || me.role === "hrManager";
  const noForms = Object.keys(HR_FORM_TYPES).length === 0;

  const newForm = (type) => { const ft = HR_FORM_TYPES[type]; if (ft) setActive({ form: ft.new(), readOnly: false }); };
  const viewForm = (r) => setActive({ form: r, readOnly: true });
  const editForm = (r) => setActive({ form: r, readOnly: false, editing: true });

  const save = async (form) => {
    setBusy(true);
    const isEdit = !!(active && active.editing);
    const saved = isEdit
      ? { ...form, editedBy: me.name, editedAt: new Date().toISOString(), approved: null, approvedBy: null }
      : { ...form, submitted: new Date().toISOString(), submittedBy: me.name, submittedUsername: me.username };
    const cur = Array.isArray(records) ? records : [];
    const next = isEdit ? cur.map((r) => r.id === saved.id ? saved : r) : [...cur, saved];
    const ok = await saveShared(HR_KEY, next);
    setRecords(next); setActive(null);
    setMsg(ok ? ["ok", isEdit ? "Record updated." : "Record submitted and saved."] : ["warn", "Saved locally but shared storage is unavailable."]);
    setBusy(false);
  };

  const approveForm = async (r) => {
    const approved = { ...r, approved: new Date().toISOString(), approvedBy: me.name, approvedUsername: me.username };
    const next = (Array.isArray(records) ? records : []).map((x) => x.id === r.id ? approved : x);
    const ok = await saveShared(HR_KEY, next);
    setRecords(next);
    if (active && active.form.id === r.id) setActive({ form: approved, readOnly: true });
    setMsg(ok ? ["ok", "Record approved."] : ["warn", "Saved locally but shared storage is unavailable."]);
  };

  const deleteHrForm = async (r) => {
    if (!window.confirm("Delete this HR record permanently? This cannot be undone.")) return;
    const next = (Array.isArray(records) ? records : []).filter((x) => x.id !== r.id);
    const ok = await saveShared(HR_KEY, next);
    setRecords(next);
    if (active && active.form.id === r.id) setActive(null);
    setMsg(ok ? ["ok", "Record deleted."] : ["warn", "Deleted locally."]);
  };

  const downloadPdf = async (r) => {
    try {
      await loadSnapLibs();
      const el = document.getElementById("hr-print-" + r.id);
      if (!el) return;
      // Force A4 portrait width (794px @ 96dpi) so content fills and centres on page
      const A4W = 794;
      el.style.width = A4W + "px";
      el.style.minWidth = A4W + "px";
      await new Promise((res) => setTimeout(res, 60));
      const canvas = await window.html2canvas(el, { scale: 2, backgroundColor: "#ffffff", useCORS: true, width: A4W, windowWidth: A4W });
      el.style.width = ""; el.style.minWidth = "";
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const M = 10, availW = 210 - 2 * M, availH = 297 - 2 * M;
      const ratio = canvas.height / canvas.width;
      const w = availW, h = availW * ratio;
      // multi-page: split tall contracts across pages
      if (h <= availH) {
        pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", M, M, w, h);
      } else {
        const pageH = Math.floor(canvas.height * (availH / h));
        let y = 0;
        while (y < canvas.height) {
          const sliceH = Math.min(pageH, canvas.height - y);
          const pageCanvas = document.createElement("canvas");
          pageCanvas.width = canvas.width; pageCanvas.height = sliceH;
          pageCanvas.getContext("2d").drawImage(canvas, 0, y, canvas.width, sliceH, 0, 0, canvas.width, sliceH);
          if (y > 0) pdf.addPage();
          const ph = sliceH * (availW / canvas.width);
          pdf.addImage(pageCanvas.toDataURL("image/jpeg", 0.92), "JPEG", M, M, w, ph);
          y += sliceH;
        }
      }
      const ft = HR_FORM_TYPES[r.type];
      pdf.save("Merrylight_HR_" + (ft ? ft.pdfName(r) : r.type + "_" + r.submitted?.slice(0,10)) + ".pdf");
    } catch (e) { setMsg(["warn", "PDF failed: " + (e && e.message)]); }
  };

  const PrintHeader = ({ r }) => {
    const ft = HR_FORM_TYPES[r.type] || {};
    return (
      <div style={{ display:"flex", alignItems:"center", gap:12, borderBottom:"2px solid #4C46A8", paddingBottom:8, marginBottom:14 }}>
        <img src={LOGO} alt="" style={{ width:36, height:36, borderRadius:"50%" }} />
        <div>
          <div style={{ fontWeight:800, fontSize:15, fontFamily:"sans-serif" }}>MERRYLIGHT ENTERPRISES — {(ft.label || "HR RECORD").toUpperCase()}</div>
          <div style={{ fontSize:11, color:"#6E6D85" }}>Submitted: {r.submitted?.slice(0,16).replace("T"," ")} · by: <b>{r.submittedBy}</b> ({r.submittedUsername || ""})</div>
          {r.approved && <div style={{ fontSize:11, color:"#1A5C2A", fontWeight:700 }}>✓ Approved by {r.approvedBy} ({r.approvedUsername || ""}) · {r.approved?.slice(0,16).replace("T"," ")}</div>}
          {r.signature && <div style={{ marginTop:4 }}><div style={{ fontSize:10, color:"#6E6D85", marginBottom:2 }}>Signature:</div><img src={r.signature} style={{ height:50, border:"1px solid #D8D9E6", borderRadius:6 }} alt="Signature" /></div>}
        </div>
      </div>
    );
  };

  if (active) {
    const ft = HR_FORM_TYPES[active.form.type] || {};
    const FormComp = ft.component;
    return (
      <>
        <div className="page-head">
          <div>
            <h2>{ft.label || active.form.type}</h2>
            <p className="muted">
              {active.readOnly ? (active.form.submitted ? "Submitted " + dfmt(active.form.submitted.slice(0,10)) + " by " + active.form.submittedBy : "Draft") : (active.editing ? "Editing" : "New record")}
              {active.form.approved ? <span className="badge ok" style={{ marginLeft:8 }}>Approved by {active.form.approvedBy}</span> : active.form.submitted ? <span className="badge" style={{ marginLeft:8, background:"#FFF3CD", color:"#856404" }}>Pending approval</span> : null}
            </p>
          </div>
          <div style={{ display:"flex", gap:6 }}>
            {active.readOnly && !active.form.approved && canEditContract && <button className="btn btn-mini" onClick={() => editForm(active.form)}>Edit</button>}
            {active.readOnly && !active.form.approved && canApprove && <button className="btn btn-primary" style={{ width:"auto", padding:"7px 12px" }} onClick={() => approveForm(active.form)}>Approve</button>}
            {active.form.approved && active.readOnly && <button className="btn btn-mini" onClick={() => downloadPdf(active.form)}>Download PDF</button>}
            {me.role === "admin" && <button className="btn btn-danger" style={{ padding:"7px 12px", width:"auto" }} onClick={() => deleteHrForm(active.form)}>Delete</button>}
            <button className="btn btn-mini" onClick={() => setActive(null)}>Back to list</button>
          </div>
        </div>
        <Card>
          {FormComp ? <FormComp form={active.form} onChange={(f) => setActive({ ...active, form: f })} onSave={save} onCancel={() => setActive(null)} me={me} readOnly={active.readOnly} /> : <p className="muted">Form not found: {active.form.type}</p>}
        </Card>
        {active.readOnly && active.form.approved && (
          <div id={"hr-print-" + active.form.id} style={{ position:"absolute", left:-9999, top:0, width:794, background:"#fff", padding:24, boxSizing:"border-box" }}>
            <PrintHeader r={active.form} />
            {FormComp && <FormComp form={active.form} onChange={() => {}} onSave={() => {}} onCancel={() => {}} me={me} readOnly={true} />}
          </div>
        )}
      </>
    );
  }

  const filtered = selType ? records.filter((r) => r && r.id && r.type === selType) : [];

  return (
    <>
      <div className="page-head">
        <div>
          <h2>{selType && HR_FORM_TYPES[selType] ? HR_FORM_TYPES[selType].label : "Human Resources"}</h2>
          <p className="muted">{selType ? filtered.length + " submission" + (filtered.length === 1 ? "" : "s") : "Select a form type to begin"}</p>
        </div>
        {canFill && !noForms && (
          <div style={{ display:"flex", gap:8 }}>
            <select className="inp" value={selType || ""} style={{ width:260 }} onChange={(e) => setSelType(e.target.value || null)}>
              <option value="">Select form type…</option>
              {Object.entries(HR_FORM_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
            {selType && <button className="btn btn-primary" style={{ width:"auto" }} onClick={() => newForm(selType)}>New form</button>}
          </div>
        )}
      </div>
      {msg && <p className={"savedmsg " + msg[0]} style={{ marginBottom:12 }}>{msg[1]}</p>}
      {noForms ? (
        <Card><p className="muted center" style={{ padding:24 }}>No HR forms configured yet. Upload your HR documents and Claude will build them here.</p></Card>
      ) : (
        <Card pad={false}>
          <div className="tbl-wrap">
            <table>
              <thead><tr><th>Date</th><th>Form</th><th>Submitted by</th><th>Status</th><th className="r">Actions</th></tr></thead>
              <tbody>
                {loading && <tr><td colSpan={5} className="muted center">Loading…</td></tr>}
                {!loading && !selType && <tr><td colSpan={5} className="muted center">Select a form type above.</td></tr>}
                {!loading && selType && !filtered.length && <tr><td colSpan={5} className="muted center">No submissions yet for {(HR_FORM_TYPES[selType] || {}).label || selType}.</td></tr>}
                {!loading && selType && filtered.sort((a,b) => (b.submitted||"").localeCompare(a.submitted||"")).map((r) => (
                  <tr key={r.id}>
                    <td>{dfmt(r.submitted?.slice(0,10))}</td>
                    <td><span className="badge">{(HR_FORM_TYPES[r.type] || {}).label?.split(" ")[0] || r.type}</span></td>
                    <td className="strong">{r.submittedBy}</td>
                    <td>{r.approved ? <span className="badge ok">Approved</span> : <span className="badge" style={{ background:"#FFF3CD",color:"#856404" }}>Pending</span>}</td>
                    <td className="r">
                      <button className="btn btn-mini" onClick={() => viewForm(r)}>View</button>{" "}
                      {!r.approved && canFill && <><button className="btn btn-mini" onClick={() => editForm(r)}>Edit</button>{" "}</>}
                      {!r.approved && canApprove && <><button className="btn btn-mini" style={{ background:"var(--bloom)",color:"#fff",border:0 }} onClick={() => approveForm(r)}>Approve</button>{" "}</>}
                      {r.approved && <><button className="btn btn-mini" onClick={() => downloadPdf(r)}>PDF</button>{" "}</>}
                      {me.role === "admin" && <button className="btn btn-mini btn-danger" style={{ padding:"3px 8px" }} onClick={() => deleteHrForm(r)}>Delete</button>}
                      <div id={"hr-print-" + r.id} style={{ position:"absolute", left:-9999, top:0, width:794, background:"#fff", padding:24, boxSizing:"border-box" }}>
                        <PrintHeader r={r} />
                        {HR_FORM_TYPES[r.type]?.component && React.createElement(HR_FORM_TYPES[r.type].component, { form: r, onChange: () => {}, onSave: () => {}, onCancel: () => {}, me, readOnly: true })}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </>
  );
}


/* ============ HR FORM: SHARED CONTRACT LAYOUT ============ */
const CONTRACT_TABLE = [
  ["Pay Date", "Last day of every month", "Sick leave p.a.", "As per statutory regulations"],
  ["Hours of work", "In accordance with the CBA / SMETA regulation", "Special Leave p.a.", "As per statutory regulations"],
  ["Accommodation", "Employer does not provide", "NSSA p.m.", "As per statutory regulations"],
  ["Fuel", "Provided to company vehicles only", "Lights", "As per statutory regulations"],
];

const cs = { fontSize: 12, lineHeight: 1.6 };

const EditContractCtx = React.createContext(false);

function ContractClause({ title, clauseKey, children, form, onChange, readOnly }) {
  const canEditContract = React.useContext(EditContractCtx);
  const override = form && clauseKey && form.customClauses && form.customClauses[clauseKey];
  const displayed = override !== undefined ? override : (typeof children === "string" ? children : null);
  const editing = !readOnly && form && clauseKey && onChange;
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3 }}>
        <div style={{ fontWeight: 700, fontSize: 12, textTransform: "uppercase", flex: 1 }}>{title}</div>
        {editing && canEditContract && <button type="button" className="btn btn-mini" style={{ padding: "1px 7px", fontSize: 10 }}
          onClick={() => {
            const cur = form.customClauses || {};
            if (cur[clauseKey] !== undefined) {
              const next = { ...cur }; delete next[clauseKey];
              onChange({ ...form, customClauses: next });
            } else {
              onChange({ ...form, customClauses: { ...cur, [clauseKey]: typeof children === "string" ? children : "" } });
            }
          }}>
          {(form.customClauses || {})[clauseKey] !== undefined ? "Reset to default" : "Edit clause"}
        </button>}
      </div>
      {editing && (form.customClauses || {})[clauseKey] !== undefined
        ? <textarea className="inp" rows={4} style={{ fontSize: 12, width: "100%", resize: "vertical" }}
            value={(form.customClauses || {})[clauseKey]}
            onChange={(e) => onChange({ ...form, customClauses: { ...(form.customClauses || {}), [clauseKey]: e.target.value } })} />
        : <div style={{ ...cs, color: "#2a2a3a", background: override !== undefined ? "#FFFBEB" : "transparent", borderRadius: 4, padding: override !== undefined ? "4px 6px" : 0 }}>
            {displayed !== null ? displayed : children}
            {override !== undefined && <span className="muted small" style={{ marginLeft: 6 }}>(customised)</span>}
          </div>}
    </div>
  );
}

function SigBlock({ label, sigKey, nameKey, dateKey, form, onChange, readOnly, canEdit }) {
  const editable = canEdit !== undefined ? canEdit : !readOnly;
  return (
    <div style={{ flex: 1, minWidth: 200 }}>
      <div style={{ fontWeight: 700, fontSize: 11, marginBottom: 6, textTransform: "uppercase" }}>{label}</div>
      <div className="fgrid" style={{ marginBottom: 6 }}>
        <label className="fld" style={{ marginBottom: 4 }}><span style={{ fontSize: 11 }}>Name</span>
          <input className="inp" style={{ fontSize: 12 }} value={form[nameKey] || ""} onChange={(e) => onChange({ ...form, [nameKey]: e.target.value })} disabled={!editable} />
        </label>
        <label className="fld" style={{ marginBottom: 4 }}><span style={{ fontSize: 11 }}>Date</span>
          <input className="inp" type="date" style={{ fontSize: 12 }} value={form[dateKey] || ""} onChange={(e) => onChange({ ...form, [dateKey]: e.target.value })} disabled={!editable} />
        </label>
      </div>
      <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 4 }}>Signature</div>
      <SignaturePad value={form[sigKey] || ""} onChange={editable ? (v) => onChange({ ...form, [sigKey]: v }) : null} disabled={!editable} />
    </div>
  );
}

function ContractSummaryTable({ noticeText }) {
  const rows = [
    ...CONTRACT_TABLE,
    ["Notice period", noticeText || "3 months' notice", "Discretionary Bonus", "At the discretion of the Employer."],
  ];
  return (
    <div className="tbl-wrap" style={{ marginBottom: 14 }}>
      <table style={{ fontSize: 11 }}>
        <tbody>
          {rows.map(([k1, v1, k2, v2], i) => (
            <tr key={i}>
              <td style={{ fontWeight: 700, background: "#F6F6FD", width: "16%", padding: "4px 8px" }}>{k1}</td>
              <td style={{ padding: "4px 8px", width: "34%" }}>{v1}</td>
              <td style={{ fontWeight: 700, background: "#F6F6FD", width: "16%", padding: "4px 8px" }}>{k2}</td>
              <td style={{ padding: "4px 8px", width: "34%" }}>{v2}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ========== CONTRACT 1: General Contract of Employment ========== */
const CONTRACT_VERSIONS = { contract: "v1.0 (SI 41 of 2022)", permcontract: "v1.0 (SI 42 of 2022)", seasonal: "v1.0 (SI 42 of 2022)", truckdriver: "v1.0 (SI 42 of 2022)" };

function newContractForm(me) {
  return {
    id: "contract-" + Date.now(), type: "contract",
    version: CONTRACT_VERSIONS.contract,
    customClauses: {},
    // Personal Information
    employeeName: "", identityNumber: "", startDate: "",
    // Job Details
    jobTitle: "", jobFunction: "See attachment", department: "",
    // Working Conditions (EXACT FROM DOCUMENT)
    daysWorked: "6 days a week", hoursPerDay: "8 hrs/day", hoursPerMonth: "208 hours per month",
    mealBreaks: "According to section management arrangement", timeOff: "One day every week",
    // Allowances & Deductions
    allowances: "NEC Agriculture stipulated allowances in terms of current Agricultural Collective Bargaining Agreement",
    nssaDeduction: "4.5% employee, 4.5% employer contribution", necDues: "1.5% employee, 1.5% employer",
    nssaDisclaimer: "Company not responsible for NSSA pension fund administration. Accept no liability for monies due upon retirement, invalidity or death.",
    // Overtime
    overtimeLimit: "Not exceeding 12 hours per week", overtimePayment: "NEC stipulated amount or time off in lieu (voluntary)",
    // Wages
    wage: "", paymentTerms: "Monthly, 4 days from end of month",
    // Protective Equipment
    protectiveClothing: "Employer property - employee responsible for maintenance",
    gumboots: "Employer property for 24 months, then becomes employee property",
    // Leave
    annualLeave: "2.5 days per month (after working 26 working days per month without fail)",
    sickLeave: "Medical certificate required",
    // Termination
    terminationNotice: "3 months notice",
    // Confidentiality
    confidentiality: "Business secrets confidential during and after employment",
    // Contract Terms
    contractTerms: "Agricultural Collective Bargaining Agreement (SI 41 of 2022)",
    // Desertation Policy
    abscondingPolicy: "Absent more than 5 days without notice = automatic termination",
    // Accommodation
    accommodation: "Merrylight does not provide. Contact HR for Springvale Farm options.",
    // Relief Work
    reliefWork: "Minimum 26 working days. Pay adjusted to grade rate.",
    // Bank Account
    bankName: "", bankBranch: "", bankAccount: "",
    // Age Certification
    ageCertification: "Employee certifies not younger than 18 years (SI 41 of 2022)",
    // Acknowledgement
    acknowledgement: "I confirm I have read this contract and accept employment conditions",
    // Signatures
    employeeSig: "", employeeSigName: "", employeeSigDate: "",
    employerSig: "", employerSigName: "", employerSigDate: "",
    witnessEmpSig: "", witnessEmpName: "", witnessEmpDate: "",
    witnessErSig: "", witnessErName: "", witnessErDate: "",
    submitted: null, submittedBy: null, submittedUsername: null,
  };
}

function ContractForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const isHRManager = me && (me.role === "admin" || me.role === "hrManager");
  const canEditFields = isHRManager && !readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  return (
    <EditContractCtx.Provider value={canEditFields}>
    <div style={cs}>
      <h3 style={{ textAlign: "center", marginBottom: 4 }}>CONTRACT OF EMPLOYMENT</h3>
      <p style={{ textAlign: "center", fontSize: 12, marginBottom: 14 }}>ENTERED INTO BETWEEN MERRYLIGHT ENTERPRISES [PRIVATE] LIMITED,<br/>121 Borrowdale Rd, Harare<br/>(Hereinafter referred to as "the employer")<br/><br/>AND</p>
      
      {/* EMPLOYEE INFO */}
      <div className="fgrid" style={{ marginBottom: 16, padding: 12, background: "#F6F6F9", borderRadius: 6 }}>
        <label className="fld"><span>Name</span><input className="inp" value={form.employeeName || ""} onChange={set("employeeName")} disabled={!canEditFields} /></label>
        <label className="fld"><span>Identity Number</span><input className="inp" value={form.idNumber || ""} onChange={set("idNumber")} disabled={!canEditFields} /></label>
      </div>
      <p style={{textAlign: "center", fontSize: 12, marginBottom: 16}}>(Hereinafter referred to as "the employee")</p>

      {/* CONTRACT CLAUSES - EXACT FROM DOCUMENT */}
      <div style={{ fontSize: 13, lineHeight: 1.6, marginBottom: 16 }}>
        <div style={{marginBottom: 14}}>
          <strong>1. DURATION</strong><br />
          Employee understands and expressly undertakes that employment duration with the employer is effective from <input className="inp" type="date" value={form.engagementDate || ""} onChange={set("engagementDate")} disabled={!canEditFields} style={{width: 130, display: "inline", padding: "2px 4px"}} />.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>2. JOB TITLE</strong><br />
          <input className="inp" value={form.jobTitle || ""} onChange={set("jobTitle")} disabled={!canEditFields} style={{width: "100%", padding: "4px"}} />
        </div>

        <div style={{marginBottom: 14}}>
          <strong>3. JOB FUNCTION</strong><br />
          See attachment
        </div>

        <div style={{marginBottom: 14}}>
          <strong>4. AGE</strong><br />
          Employee certifies that he/she is not younger than 18 years in terms of Statutory Instrument 41 of 2022
        </div>

        <div style={{marginBottom: 14}}>
          <strong>5. ALLOWANCES</strong><br />
          MERRYLIGHT ENTERPRISES [PVT] LTD provides the NEC Agriculture stipulated allowances in terms of the current Agricultural Collective Bargaining Agreement issued by NEC.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>6. DAYS TO BE WORKED</strong><br />
          Farm worker shall work the following days of the week:<br />
          • 6.1.1: 6 days a week<br />
          • 6.1.2: 8 hrs/day
        </div>

        <div style={{marginBottom: 14}}>
          <strong>7. HOURS OF WORK</strong><br />
          • Ordinary hours of work shall not exceed <strong>208 hours per month</strong><br />
          • Employee must conform to company hours of work and days to be worked per week<br />
          • Employee expected to work additional hours as required by management for effective performance of duties<br />
          • Coming to work late and leaving early will not be acceptable - Code of Conduct will be instituted
        </div>

        <div style={{marginBottom: 14}}>
          <strong>8. MEAL BREAKS</strong><br />
          Breaks in between working hours according to each section management arrangement.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>9. TIME OFF</strong><br />
          Employee entitled to: <strong>One day every week</strong>
        </div>

        <div style={{marginBottom: 14}}>
          <strong>10. OVERTIME</strong><br />
          • Employer may ask employee to volunteer to work overtime<br />
          • <strong>Approved limits:</strong> Not exceeding 12 hours per week<br />
          • <strong>Payment:</strong> NEC stipulated amount as required by law<br />
          • <strong>Options:</strong> Paid at NEC rates OR time off in lieu of overtime worked<br />
          • <strong>Note:</strong> Overtime is worked on voluntary basis
        </div>

        <div style={{marginBottom: 14}}>
          <strong>11. WAGES</strong><br />
          <strong>11.1</strong> Fixed wage as agreed by management and will change from time to time according to performance.<br /><br />
          <strong>11.2</strong> Wage: <input className="inp" value={form.wage || ""} onChange={set("wage")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} /><br /><br />
          <strong>11.3</strong> Payment: Monthly, 4 days from end of month
        </div>

        <div style={{marginBottom: 14}}>
          <strong>12. NSSA AND DUES</strong><br />
          • <strong>NSSA Deduction:</strong> 4.5% of wage deducted for NSSA<br />
          • <strong>NSSA Contribution:</strong> Company contributes 4.5% on employee's behalf<br />
          • <strong>NEC Agriculture Dues:</strong> Company deducts 1.5%<br />
          • <strong>NEC Dues Contribution:</strong> Company contributes 1.5% on employee's behalf<br />
          • <strong>Effective Date:</strong> First day of employment<br />
          • <strong>Liability Disclaimer:</strong> Company not responsible for NSSA pension fund administration and accepts no liability for monies due from NSSA fund upon retirement, invalidity or untimely death
        </div>

        <div style={{marginBottom: 14}}>
          <strong>13. PROTECTIVE CLOTHING</strong><br />
          Protective clothing supplied shall remain the property of the employer. Employee is responsible for mending, washing and otherwise maintaining such clothing.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>14. GUMBOOTS</strong><br />
          Gumboots shall remain the property of the employer for a period of 24 months, thereafter it becomes the employee's property.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>15. ANNUAL LEAVE</strong><br />
          • Employee accumulates paid vacation leave at rate of <strong>2.5 days only after working full month (26 working days)</strong><br />
          • <strong>Condition:</strong> Must work 26 working days each month without fail<br />
          • <strong>Penalty:</strong> Failure to work 26 working days due to absenteeism - company will not pay leave for that month
        </div>

        <div style={{marginBottom: 14}}>
          <strong>16. SICK LEAVE</strong><br />
          Sick leave shall only be granted at request of employee supported by certificate signed by medical practitioner granting the employee the days booked as off-duty.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>17. TERMINATION</strong><br />
          During currency of contract, parties may terminate by giving the other party: <strong>3 months' notice</strong>
        </div>

        <div style={{marginBottom: 14}}>
          <strong>18. CONFIDENTIAL DOCUMENTS</strong><br />
          Employee will consider as business secrets all information acquired regarding company matters. At no time during term of contract or subsequently may such information be divulged in any way whatsoever to any third party without prior written consent of the company.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>19. CONDITIONS OF THE CONTRACT</strong><br />
          This contract is in terms of the Agricultural Collective Bargaining Agreement (SI 41 of 2022) as amended from time to time. Disciplinary hearing procedures and termination of contract shall be in terms of said regulations.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>20. DEEMED DESERTATION</strong><br />
          • Employee must immediately inform employer if absent from work for any reason<br />
          • <strong>Automatic Termination Clause:</strong> If absent for more than <strong>5 days without having reported or informed employer</strong> of reasons for absence, farm worker will be deemed to have absconded<br />
          • <strong>Consequence:</strong> Services will be terminated automatically
        </div>

        <div style={{marginBottom: 14}}>
          <strong>21. ACCOMMODATION</strong><br />
          Merrylight does not have any accommodation. If you would like accommodation at Springvale Farm, please advise HR so they can liaise with the farm owner.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>22. RELIEF WORK</strong><br />
          • Employer may request relief work in higher grade than normally employed<br />
          • <strong>Minimum Duration:</strong> Continuous period of not less than 26 working days<br />
          • <strong>Pay Adjustment:</strong> Employer shall pay wage adjustment from date started in higher grade<br />
          • <strong>Wage Guarantee:</strong> Daily wage shall be not less than wage applicable to such grade<br />
          • <strong>Reversion:</strong> Once task completed, employee required to revert to original grade and pay
        </div>

        <div style={{marginBottom: 14}}>
          <strong>23. BANK ACCOUNT DETAILS</strong><br />
          I hereby state that my salary should be paid into the below stipulated account:<br /><br />
          <strong>Bank:</strong> <input className="inp" value={form.bankName || ""} onChange={set("bankName")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} /><br />
          <strong>Branch:</strong> <input className="inp" value={form.bankBranch || ""} onChange={set("bankBranch")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} /><br />
          <strong>Account Number:</strong> <input className="inp" value={form.bankAccount || ""} onChange={set("bankAccount")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} />
        </div>

        <div style={{marginBottom: 16}}>
          <strong>24. ACKNOWLEDGEMENT & SIGNATURE</strong><br />
          I hereby confirm that I have read the attached contract document in full and confirm that this reflects the full offer of employment. I accept the conditions of employment as detailed on this contract.
        </div>
      </div>

      {/* SIGNATURES */}
      <div style={{ display: "flex", gap: 24, flexWrap: "wrap", marginBottom: 16 }}>
        <SigBlock label="Employer" sigKey="employerSig" nameKey="employerSigName" dateKey="employerSigDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
        <SigBlock label="Employee" sigKey="employeeSig" nameKey="employeeSigName" dateKey="employeeSigDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
      </div>
      <div style={{ display: "flex", gap: 24, flexWrap: "wrap", marginBottom: 16 }}>
        <SigBlock label="Witness (Employer's)" sigKey="witnessErSig" nameKey="witnessErName" dateKey="witnessErDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
        <SigBlock label="Witness (Employee's)" sigKey="witnessEmpSig" nameKey="witnessEmpName" dateKey="witnessEmpDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
      </div>

      {!readOnly && <div style={{ display: "flex", gap: 8, marginTop: 14 }}><button className="btn btn-primary" style={{ width: "auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
    </EditContractCtx.Provider>
  );
}

/* ========== CONTRACT 2: Permanent Contract of Employment ========== */
function newPermanentContractForm(me) {
  return {
    id: "permcontract-" + Date.now(), type: "permcontract",
    version: CONTRACT_VERSIONS.permcontract,
    customClauses: {},
    // Personal Information
    surname: "", firstName: "", dateOfBirth: "", idNumber: "",
    // Job Details
    jobTitle: "", grade: "", farmName: "",
    // Engagement
    engagementDate: new Date().toISOString().slice(0, 10),
    duration: "PERMANENT",
    // Wages & Benefits (EXACT FROM DOCUMENT)
    wageAmount: "Stipulated by NEC for Horticulture", allowances: "Transport, accommodation, fuel, lights, dog handling, firearm, travel and subsistence per NEC",
    paymentMethod: "Electronic transfer monthly",
    // Working Schedule
    hoursPerWeek: "Minimum 48 hours per week",
    farmOperations: "6 days/week 06:30-15:00 or 07:30-16:00 with 1/2 hour tea break",
    guards: "4 days weekday SHIFT 06:00-18:00, NIGHT SHIFT 18:00-06:00, 5th day overtime rates",
    drivers: "4 days weekday SHIFT 04:30am-16:30",
    hoursAmendment: "May be amended at Management discretion but not exceed 48 hours normal",
    // Government Holidays
    governmentHolidays: "Paid if not required to work. If required, paid as per regulations.",
    // Transport
    transport: "Supplied from specific pickup points if not within walking distance",
    // Hygiene & Safety
    hygienePolicy: "Employee must adhere to hygiene, food safety code and workplace procedures per Zimbabwe laws",
    // Children
    childrenPolicy: "Children not allowed at workplace. Babies can stay at crèche provided.",
    // Overtime (EXACT FROM DOCUMENT)
    overtimeLimit: "Not exceeding 12 hours per week", overtimePayment: "NEC stipulated amount or time off in lieu (voluntary)",
    peakSeasonTransport: "Truck may leave late if workers in overtime during peak seasons. Non-participating transported.",
    // Protective Equipment
    protectiveClothing: "Employer property - employee responsible for maintenance",
    gumboots: "Employer property for 24 months, then becomes employee property",
    // Annual Leave
    annualLeave: "0.096 days per day worked (2.5 days per month) - paid monthly",
    // Sick Leave
    sickLeave: "Medical certificate from practitioner required",
    // Compensation Summary
    payDate: "Last day of every month", sickLeavePA: "As per statutory regulations",
    hoursOfWork: "Per CBA & SMETA regulation", specialLeavePA: "As per statutory regulations",
    accommodation: "Not provided", nssaDeduction: "Per statutory regulations",
    fuel: "Provided to company vehicles only", lights: "Per statutory regulations",
    noticeRequired: "3 months notice", discretionaryBonus: "At discretion of Employer",
    // Termination
    terminationNotice: "3 months notice",
    // Confidentiality
    confidentiality: "Not divulge confidential information during or after employment",
    confidentialityException: "Except when auditing body or Zimbabwe laws require disclosure",
    // Contract Terms
    contractTerms: "Agricultural Collective Bargaining Agreement (SI 42 OF 2022)",
    // Desertation
    abscondingPolicy: "Absent more than 5 days without reporting = deemed absconded, automatic termination",
    // Accommodation
    accommodationOffered: "Not provided. Contact HR for Springvale Farm options.",
    // Pregnancy
    pregnancyPolicy: "Female employees must advise management so appropriate duties appointed",
    // ETI Compliance
    etiPolicy: "Committed to ETI base code international labour practices. Employee must read and understand before signing.",
    // Relief Work
    reliefWork: "May request relief in higher grade for minimum 26 working days. Wage adjusted. Daily wage >= applicable grade rate.",
    // Signatures
    employeeSig: "", employeeSigName: "", employeeSigDate: "",
    employerSig: "", employerSigName: "", employerSigDate: "",
    witnessSig: "", witnessName: "", witnessDate: "",
    submitted: null, submittedBy: null, submittedUsername: null,
  };
}

function PermanentContractForm({ form, onChange, onSave, onCancel, me, readOnly }) {
  const isHRManager = me && (me.role === "admin" || me.role === "hrManager");
  const canEditFields = isHRManager && !readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  return (
    <EditContractCtx.Provider value={canEditFields}>
    <div style={cs}>
      <h3 style={{ textAlign: "center", marginBottom: 4 }}>PERMANENT CONTRACT OF EMPLOYMENT</h3>
      <p style={{ textAlign: "center", fontSize: 12, marginBottom: 14, fontWeight: 600 }}>PARTIES</p>
      <p style={{ textAlign: "center", fontSize: 12, marginBottom: 14 }}>Memorandum of an Agreement entered between:</p>
      <p style={{ textAlign: "center", fontSize: 12, marginBottom: 14 }}>MERRYLIGHT<br />(Hereinafter called the "Contract Employer" of one part)<br /><br />And</p>
      
      {/* EMPLOYEE INFO */}
      <div className="fgrid" style={{ marginBottom: 16, padding: 12, background: "#F6F6F9", borderRadius: 6 }}>
        <label className="fld"><span>Surname</span><input className="inp" value={form.surname || ""} onChange={set("surname")} disabled={!canEditFields} /></label>
        <label className="fld"><span>First Name</span><input className="inp" value={form.firstName || ""} onChange={set("firstName")} disabled={!canEditFields} /></label>
        <label className="fld"><span>Date of Birth</span><input className="inp" type="date" value={form.dateOfBirth || ""} onChange={set("dateOfBirth")} disabled={!canEditFields} /></label>
        <label className="fld"><span>ID. No</span><input className="inp" value={form.idNumber || ""} onChange={set("idNumber")} disabled={!canEditFields} /></label>
      </div>

      {/* CONTRACT CLAUSES - EXACT FROM DOCUMENT */}
      <div style={{ fontSize: 13, lineHeight: 1.6, marginBottom: 16 }}>
        <div style={{marginBottom: 14}}>
          <strong>1. DURATION</strong><br />
          IT IS HEREBY AGREED THAT: The Contract Employer shall employ the Assignee as a <input className="inp" value={form.jobTitle || ""} onChange={set("jobTitle")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} /> at job grade <input className="inp" value={form.grade || ""} onChange={set("grade")} disabled={!canEditFields} style={{width: 100, display: "inline", padding: "2px 4px"}} /> stationed at farm name <input className="inp" value={form.farmName || ""} onChange={set("farmName")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} />. By entering this contract, employee understands and expressly undertakes that duration of employment is PERMANENT with effect from <input className="inp" type="date" value={form.engagementDate || ""} onChange={set("engagementDate")} disabled={!canEditFields} style={{width: 130, display: "inline", padding: "2px 4px"}} />.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>2. WAGE & PAYMENT</strong><br />
          • Grade calculated as per current Collective Bargaining Agreement (CBA) issued by NEC<br />
          • Wages paid: Monthly<br />
          • Wage amount: Calculated at amount stipulated by NEC for Horticulture<br />
          • Allowances: Transport, accommodation, fuel, lights, dog handling, firearm or travel and subsistence - paid as per NEC stipulated amount and pro-rated<br />
          • Latest Wage Document: Available at HR<br />
          • Bonus Schemes: Management may introduce various incentives and/or bonus schemes - details will be communicated at time of introducing/withdrawing
        </div>

        <div style={{marginBottom: 14}}>
          <strong>3. DAYS & HOURS TO BE WORKED</strong><br />
          Employee expected to work <strong>minimum of 48 hours per week</strong><br /><br />
          <strong>Merrylight Farm Operations:</strong><br />
          • 6 days/week<br />
          • 06:30 to 15:00 OR 07:30 to 16:00<br />
          • ½ hour tea break<br /><br />
          <strong>Merrylight Guards:</strong><br />
          • 4 days/weekday SHIFT: 06:00 to 18:00 (depending on duty)<br />
          • NIGHT SHIFT: 18:00 to 06:00 (depending on duty)<br />
          • Fifth day worked available at overtime rates<br /><br />
          <strong>Drivers:</strong><br />
          • 4 days/weekday SHIFT: 04:30 AM to 16:30 (depending on duty)<br /><br />
          Amendment: Hours may be amended at Management's discretion but should not exceed 48 hours normal hours per week
        </div>

        <div style={{marginBottom: 14}}>
          <strong>4. GOVERNMENT HOLIDAYS</strong><br />
          Employee will be paid for all agricultural holidays as advised by government. If required to work, payment as per NEC regulations.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>5. MEAL BREAKS</strong><br />
          There shall be breaks between working hours according to each section management arrangement.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>6. TRANSPORT</strong><br />
          Transport supplied by company from specific pickup points, providing you do not live within walking distance.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>7. HYGIENE & FOOD SAFETY</strong><br />
          Employee must adhere to employer's hygiene food safety code and any other procedures or policy the employer stipulates within the laws of Zimbabwe.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>8. CHILDREN AT WORKPLACE</strong><br />
          • Employer does not allow children to be present at workplace<br />
          • Babies can be brought but must stay at crèche provided
        </div>

        <div style={{marginBottom: 14}}>
          <strong>9. OVERTIME</strong><br />
          • Employer may ask employee to volunteer to work overtime<br />
          • Approved limits: Not exceeding 12 hours per week<br />
          • Payment: NEC stipulated amount or as required by law<br />
          • Options: Paid at NEC rates OR time off in lieu of overtime worked<br />
          • Note: Overtime is voluntary basis<br />
          • Peak Season Transport: Employer may require truck to leave facilities late if workers involved in overtime (peak seasons only). Non-participating employees will be transported.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>10. PROTECTIVE CLOTHING</strong><br />
          Protective clothing supplied shall remain property of employer. Employee responsible for mending, washing and otherwise maintaining such clothing.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>11. GUMBOOTS</strong><br />
          Shall remain property of employer for period of 24 months, thereafter becomes employee's property.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>12. ANNUAL LEAVE</strong><br />
          Calculated at rate of <strong>0.096 days per day worked (2.5 days per month)</strong>. This will be paid monthly.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>13. SICK LEAVE</strong><br />
          Sick leave shall only be granted at request of employee supported by certificate signed by medical practitioner granting days as off-duty.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>14. COMPENSATION & BENEFITS SUMMARY</strong><br />
          <table style={{width: "100%", borderCollapse: "collapse", marginTop: 8, fontSize: 12}}>
            <thead>
              <tr style={{background: "#E8E8F0"}}>
                <th style={{border: "1px solid #D0D0D8", padding: 8, textAlign: "left"}}>Item</th>
                <th style={{border: "1px solid #D0D0D8", padding: 8, textAlign: "left"}}>Details</th>
              </tr>
            </thead>
            <tbody>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Pay Date</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>Last day of every month</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Sick leave p.a.</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Hours of work</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>In accordance with CBA & SMETA regulation</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Special Leave p.a.</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Accommodation</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>Employer does not provide</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>NSSA p.m.</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Fuel</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>Provided to company vehicles only</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Lights</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Notice period</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>3 months notice</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Discretionary Bonus</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>At discretion of Employer</td></tr>
            </tbody>
          </table>
        </div>

        <div style={{marginBottom: 14}}>
          <strong>15. TERMINATION</strong><br />
          During currency of contract, parties may terminate by giving other party: <strong>3 months' notice</strong>
        </div>

        <div style={{marginBottom: 14}}>
          <strong>16. CONFIDENTIAL DOCUMENTS</strong><br />
          Assignee shall not, either during or after expiration of contract and/or period of service with Contract Employer, divulge any confidential information relating to affairs, finances or any other official matter pertaining to Client or Contract Employer to any person within or outside Zimbabwe unless with express approval of Contract Employer or Client. Company undertakes to keep all employee details confidential except when requested by auditing body or as required by laws of Zimbabwe.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>17. CONDITIONS OF THE CONTRACT</strong><br />
          This contract is in terms of Agricultural Collective Bargaining Agreement (SI 42 OF 2022) as amended from time to time. Disciplinary hearing procedures and termination of contract shall be in terms of said regulations.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>18. DEEMED DESERTATION</strong><br />
          • Employee must immediately inform employer if absent from work<br />
          • Automatic Termination: If absent for more than <strong>5 days without reporting or informing employer</strong> of reasons for absence, farm worker will be deemed to have absconded<br />
          • Consequence: Services will be terminated automatically
        </div>

        <div style={{marginBottom: 14}}>
          <strong>19. ACCOMMODATION</strong><br />
          Merrylight does not have any accommodation. If you would like accommodation at Springvale Farm, please advise HR so they can liaise with farm owner.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>20. PREGNANCY</strong><br />
          Female employees must advise management of pregnancy so that appropriate duties can be appointed.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>21. ETI BASE CODE COMPLIANCE</strong><br />
          Merrylight Enterprises is committed to complying with international best practice labour practices through ETI base code. Please read base code and ensure you understand it before signing contract.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>22. RELIEF WORK</strong><br />
          • Employer may request relief work in higher grade<br />
          • Minimum Duration: Not less than 26 working days<br />
          • Pay Adjustment: Employer shall pay wage adjustment from date started in higher grade<br />
          • Wage Guarantee: Daily wage shall be not less than wage applicable to such grade<br />
          • Reversion: Once task completed, required to revert to original grade and pay
        </div>

        <div style={{marginBottom: 14}}>
          <strong>23. WHOLE AGREEMENT</strong><br />
          This agreement constitutes the whole agreement between parties and no warranties or representations whether express or implied not stated herein shall be binding on parties. Variations to terms and conditions of this agreement must be in writing and signed by both parties.
        </div>

        <div style={{marginBottom: 16}}>
          <strong>24. SIGNATURES</strong><br />
          Engagement date: <input className="inp" type="date" value={form.engagementDate || ""} onChange={set("engagementDate")} disabled={!canEditFields} style={{width: 130, display: "inline", padding: "2px 4px"}} />
        </div>
      </div>

      {/* SIGNATURES */}
      <div style={{ display: "flex", gap: 24, flexWrap: "wrap", marginBottom: 16 }}>
        <SigBlock label="Employee" sigKey="employeeSig" nameKey="employeeSigName" dateKey="employeeSigDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
        <SigBlock label="Employer" sigKey="employerSig" nameKey="employerSigName" dateKey="employerSigDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
      </div>
      <SigBlock label="Witness" sigKey="witnessSig" nameKey="witnessName" dateKey="witnessDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
      <p style={{fontSize: 11, color: "#666", marginTop: 12}}>NOTE: The witness should be conversant in English and the mother language of the employee and should confirm the employee understands the content.</p>

      {!readOnly && <div style={{ display: "flex", gap: 8, marginTop: 14 }}><button className="btn btn-primary" style={{ width: "auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
    </EditContractCtx.Provider>
  );
}

/* ========== CONTRACT 3: Seasonal Contract ========== */
function newSeasonalContractForm(me) {
  return {
    id: "seasonal-" + Date.now(), type: "seasonal",
    version: CONTRACT_VERSIONS.seasonal,
    customClauses: {},
    // Personal Information
    surname: "", firstName: "",
    // Job Details
    grade: "SEASONAL CONTRACT worker", occupation: "", department: "",
    // Engagement Period
    engagementStart: "", engagementEnd: "",
    // Engagement Terms (EXACT FROM DOCUMENT)
    gradeNote: "Calculated per current CBA issued by NEC",
    // Wages & Payment
    wageAmount: "Stipulated by NEC for Horticulture", allowances: "Transport, Accommodation, fuel, lights, dog handling, firearm or travel and subsistence per NEC",
    paymentMethod: "Electronic transfer monthly (USD and ZIG account)",
    // Working Schedule
    workingPattern: "Maximum consecutive 6-day week with one free day per week",
    hoursPerWeek: "Minimum 48 hours per week",
    farmOperations: "6 days/week 06:30-15:00 or 07:30-16:00 with 1/2 hour tea break",
    guards: "4 days weekday SHIFT 06:00-18:00, NIGHT SHIFT 18:00-06:00, 5th day overtime rates",
    hoursAmendment: "May be amended at Management discretion but not exceed 48 hours normal",
    // Accommodation
    accommodation: "Not provided. Contact HR for Springvale Farm options.",
    // Transport
    transport: "Supplied from specific pickup points if not within walking distance. At employer's risk. 15 minutes before starting time.",
    // Termination Notice (DIFFERENT FROM PERMANENT)
    terminationNoticeEarly: "1 day notice (first 7 days of employment)",
    terminationNoticeLater: "2 weeks notice (after 7 days)",
    // Government Holidays
    governmentHolidays: "Paid as advised by government. If required to work, paid at 2.5 rate.",
    // Meal Breaks
    mealBreaks: "Breaks between working hours per section management arrangement",
    // Overtime (DIFFERENT: 10 HOURS NOT 12)
    overtimeLimit: "Not exceeding 10 hours per week", overtimePayment: "NEC stipulated amount or time off in lieu (voluntary)",
    // Bonus & Incentives
    bonusSchemes: "Management may introduce various incentives/bonus schemes - details communicated at time",
    // Hygiene & Safety
    hygienePolicy: "Employee must adhere to hygiene, food safety code and workplace procedures per Zimbabwe laws",
    // Children
    childrenPolicy: "Children not allowed at workplace. Babies can stay at break areas provided.",
    // Leave & Benefits
    payDate: "Last day of every month", sickLeavePA: "As per statutory regulations",
    hoursOfWork: "Per CBA & SMETA regulation", specialLeavePA: "As per statutory regulations",
    nssaDeduction: "Per statutory regulations", fuel: "Provided to company vehicles only",
    lights: "Per statutory regulations", noticeRequired: "1 days' notice", discretionaryBonus: "At discretion of Employer",
    // Annual Leave
    annualLeave: "0.096 days per day worked (2.5 days per month) - paid monthly",
    // Absenteeism
    absenteeismPolicy: "Anyone absent without advising management for 5 days will be dismissed per Code of Conduct",
    // Pregnancy
    pregnancyPolicy: "Female employees must advise management so appropriate duties appointed",
    // ETI Compliance
    etiPolicy: "Committed to ETI base code international labour practices. Employee must read and understand before signing.",
    // Deductions
    deductions: "Per NEC stipulation and local government",
    // Contract Renewal
    renewalPolicy: "Contract Employer may elect to renew or offer new contract. New contract NOT extension, creates no expectation of renewal or permanent employment.",
    // Confidentiality
    confidentiality: "Not divulge confidential information during or after employment",
    confidentialityNote: "PERSONAL INFO KEPT CONFIDENTIAL BUT SHARED WITH AUDITING AND CUSTOMERS WHEN REQUIRED INCLUDING PAY AND ATTENDANCE",
    // Truck Driver Specific (added dynamically if isTruckDriver)
    jamportClause: "",
    // Signatures
    employeeSig: "", employeeSigName: "", employeeSigDate: "",
    employerSig: "", employerSigName: "", employerSigDate: "",
    witnessSig: "", witnessName: "", witnessDate: "",
    submitted: null, submittedBy: null, submittedUsername: null,
  };
}

function SeasonalContractForm({ form, onChange, onSave, onCancel, me, readOnly, isTruckDriver }) {
  const isHRManager = me && (me.role === "admin" || me.role === "hrManager");
  const canEditFields = isHRManager && !readOnly;
  const set = (k) => (e) => onChange({ ...form, [k]: e.target.value });
  const title = isTruckDriver ? "SEASONAL CONTRACT OF EMPLOYMENT — TRUCK DRIVERS" : "SEASONAL CONTRACT OF EMPLOYMENT";
  return (
    <EditContractCtx.Provider value={canEditFields}>
    <div style={cs}>
      <h3 style={{ textAlign: "center", marginBottom: 4 }}>{title}</h3>
      <p style={{ textAlign: "center", fontSize: 12, marginBottom: 14, fontWeight: 600 }}>PARTIES</p>
      <p style={{ textAlign: "center", fontSize: 12, marginBottom: 14 }}>Memorandum of an Agreement entered between:<br /><br />MERRYLIGHT<br />(Hereinafter called the "Contract Employer" of one part)<br /><br />And</p>
      
      {/* EMPLOYEE INFO */}
      <div className="fgrid" style={{ marginBottom: 16, padding: 12, background: "#F6F6F9", borderRadius: 6 }}>
        <label className="fld"><span>Surname</span><input className="inp" value={form.surname || ""} onChange={set("surname")} disabled={!canEditFields} /></label>
        <label className="fld"><span>First Names</span><input className="inp" value={form.firstName || ""} onChange={set("firstName")} disabled={!canEditFields} /></label>
      </div>

      {/* CONTRACT CLAUSES - EXACT FROM DOCUMENT */}
      <div style={{ fontSize: 13, lineHeight: 1.6, marginBottom: 16 }}>
        <div style={{marginBottom: 14}}>
          <strong>1. ENGAGEMENT & ROLE</strong><br />
          Employee is engaged as a: <strong>SEASONAL CONTRACT worker</strong><br /><br />
          <strong>Grade:</strong> <input className="inp" value={form.grade || ""} onChange={set("grade")} disabled={!canEditFields} style={{width: 100, display: "inline", padding: "2px 4px"}} /><br />
          <strong>Main duties:</strong> <input className="inp" value={form.occupation || ""} onChange={set("occupation")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} /> in the <input className="inp" value={form.department || ""} onChange={set("department")} disabled={!canEditFields} style={{width: 150, display: "inline", padding: "2px 4px"}} /> department but may be given other tasks from time to time<br /><br />
          <strong>Note:</strong> Grade calculated as per current Collective Bargaining Agreement (CBA) issued by NEC.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>2. WAGES & PAYMENT</strong><br />
          • Wages paid: Monthly<br />
          • <strong>Wage amount:</strong> Calculated at amount stipulated by NEC for Horticulture<br />
          • <strong>Latest Wage Document:</strong> Available at HR<br />
          • <strong>Allowances:</strong> Transport, Accommodation, fuel lights, dog handling, firearm or travel and subsistence - paid as per NEC stipulated amount and pro-rated<br />
          • <strong>NEC Allowances:</strong> Available at HR office<br />
          • <strong>Payment Method:</strong> Electronic transfer from company to employee's stipulated USD and ZIG account (entered on new employee form)
        </div>

        <div style={{marginBottom: 14}}>
          <strong>3. WORKING SCHEDULE</strong><br />
          • Employee required to work: <strong>Maximum consecutive 6-day week with one free day per week</strong><br />
          • Expected to work: <strong>Minimum 48 hours per week</strong><br /><br />
          <strong>Merrylight Farm Operations:</strong><br />
          • 6 days/week<br />
          • 06:30 to 15:00 OR 07:30 to 16:00<br />
          • ½ hour tea break<br /><br />
          <strong>Merrylight Guards:</strong><br />
          • 4 days/weekday SHIFT: 06:00 to 18:00 (depending on duty)<br />
          • NIGHT SHIFT: 18:00 to 06:00 (depending on duty)<br />
          • Fifth day worked available at overtime rates<br /><br />
          <strong>Hours Amendment:</strong> May be amended at Management's discretion but should not exceed 48 hours normal hours per week
        </div>

        <div style={{marginBottom: 14}}>
          <strong>4. ACCOMMODATION</strong><br />
          Merrylight does not have any accommodation. If you would like accommodation at Springvale Farm, please advise HR so they can liaise with farm owner.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>5. TRANSPORT</strong><br />
          • Transport supplied by company from specific pickup points, providing you do not live within walking distance<br />
          • Travel on service provided at employer's risk<br />
          • Times: 15 minutes before starting time
        </div>

        <div style={{marginBottom: 14}}>
          <strong>6. TERMINATION NOTICE</strong><br />
          • <strong>First 7 days:</strong> One day's notice<br />
          • <strong>After 7 days:</strong> Two weeks' notice
        </div>

        <div style={{marginBottom: 14}}>
          <strong>7. GOVERNMENT HOLIDAYS</strong><br />
          Employee will be paid for all agricultural holidays as advised by government. If required to work, it will be paid at 2.5 (NEC rate).
        </div>

        <div style={{marginBottom: 14}}>
          <strong>8. MEAL BREAKS</strong><br />
          There shall be breaks between working hours according to each section management arrangement.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>9. OVERTIME</strong><br />
          • Employer may ask employee to volunteer to work overtime<br />
          • <strong>Approved limits:</strong> Not exceeding 10 hours per week<br />
          • <strong>Payment:</strong> NEC stipulated amount or as required by law<br />
          • <strong>Options:</strong> Paid at NEC rates OR time off in lieu of overtime worked<br />
          • <strong>Note:</strong> Overtime is voluntary basis
        </div>

        <div style={{marginBottom: 14}}>
          <strong>10. BONUS & INCENTIVE SCHEMES</strong><br />
          Management may introduce various incentives and/or bonus schemes. Details will be communicated at time of introducing/withdrawing such scheme.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>11. HYGIENE & FOOD SAFETY</strong><br />
          Employee must adhere to employer's hygiene food safety code and any other procedures or policy the employer stipulates within the laws of Zimbabwe.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>12. CHILDREN AT WORKPLACE</strong><br />
          • Employer does not allow children to be present at workplace<br />
          • Babies can be brought but need to stay at break areas provided
        </div>

        <div style={{marginBottom: 14}}>
          <strong>13. LEAVE & BENEFITS SUMMARY</strong><br />
          <table style={{width: "100%", borderCollapse: "collapse", marginTop: 8, fontSize: 12}}>
            <thead>
              <tr style={{background: "#E8E8F0"}}>
                <th style={{border: "1px solid #D0D0D8", padding: 8, textAlign: "left"}}>Item</th>
                <th style={{border: "1px solid #D0D0D8", padding: 8, textAlign: "left"}}>Details</th>
              </tr>
            </thead>
            <tbody>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Pay Date</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>Last day of every month</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Sick leave p.a.</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Hours of work</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>In accordance with CBA & SMETA regulation</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Special Leave p.a.</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Accommodation</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>Employer does not provide</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>NSSA p.m.</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Fuel</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>Provided to company vehicles only</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Lights</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>As per statutory regulations</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Notice period</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>1 days' notice</td></tr>
              <tr><td style={{border: "1px solid #D0D0D8", padding: 8}}>Discretionary Bonus</td><td style={{border: "1px solid #D0D0D8", padding: 8}}>At discretion of Employer</td></tr>
            </tbody>
          </table>
        </div>

        <div style={{marginBottom: 14}}>
          <strong>14. ANNUAL LEAVE</strong><br />
          Calculated at rate of <strong>0.096 days per day worked (2.5 days per month)</strong>. This will be paid monthly.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>15. ABSENTEEISM</strong><br />
          • Anyone absent without advising management for <strong>five days will be dismissed</strong> according to Code of Conduct
        </div>

        <div style={{marginBottom: 14}}>
          <strong>16. PREGNANCY</strong><br />
          Female employees must advise management of pregnancy so that appropriate duties can be appointed.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>17. ETI BASE CODE COMPLIANCE</strong><br />
          Merrylight Enterprises is committed to complying with international best practice labour practices through ETI base code. Please read base code and ensure you understand it before signing contract.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>18. DEDUCTIONS</strong><br />
          Deductions will be done according to NEC stipulation and local government.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>19. RENEWAL OF CONTRACT</strong><br />
          Contract Employer reserves right to elect whether to renew Assignee's contract or offer new contract upon expiration. If offered further contract at expiration, shall be subject to terms and conditions of new and separate agreement given at Contract Employer's entire discretion and shall NOT:<br />
          • Be regarded as extension of present agreement<br />
          • Create expectation that at expiry of fixed term contract you will be engaged on new and further contract<br />
          • Create expectation that employment shall be on permanent basis
        </div>

        <div style={{marginBottom: 14}}>
          <strong>20. CONFIDENTIALITY</strong><br />
          Assignee shall not, either during or after expiration of contract and/or period of service with Contract Employer, divulge any confidential information relating to affairs, finances or any other official matter pertaining to Client or Contract Employer to any person within or outside Zimbabwe unless with express approval of Contract Employer or Client.<br /><br />
          <strong>NOTE:</strong> PERSONAL INFORMATION SHALL BE KEPT CONFIDENTIAL BUT SHALL BE SHARED WITH AUDITING AND CUSTOMERS WHEN REQUIRED INCLUDING PAY AND ATTENDANCE.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>21. WHOLE AGREEMENT</strong><br />
          This agreement constitutes the whole agreement between parties and no warranties or representations whether express or implied not stated herein shall be binding on parties. Variations to terms and conditions of this agreement must be in writing and signed by both parties.
        </div>

        <div style={{marginBottom: 14}}>
          <strong>22. SIGNATURES</strong><br />
          Engagement Date: <input className="inp" type="date" value={form.engagementStart || ""} onChange={set("engagementStart")} disabled={!canEditFields} style={{width: 130, display: "inline", padding: "2px 4px"}} /> <strong>to</strong> <input className="inp" type="date" value={form.engagementEnd || ""} onChange={set("engagementEnd")} disabled={!canEditFields} style={{width: 130, display: "inline", padding: "2px 4px"}} />
        </div>

        {isTruckDriver && (
          <div style={{marginBottom: 14, padding: 12, background: "#FFF8E1", borderRadius: 6}}>
            <strong>SPECIAL CLAUSE 19: HOURS OF WORK - TRANSPORT DRIVERS (JAMPORT)</strong><br />
            Employees engaged as transport drivers through Jamport shall be regarded as employees of Jamport while transporting employees from designated pick-up points to Merrylight Enterprises at commencement of working day. Upon arrival at Merrylight Enterprises and after parking their vehicles, the drivers shall remain at the workplace until the scheduled employee collection time. During this period, the drivers may, at their own discretion, elect to perform driving or other assigned duties for Merrylight Enterprises. Such work shall be entirely voluntary. No driver shall be compelled, coerced, or disadvantaged for choosing not to undertake such duties. Where a driver voluntarily agrees to perform duties for Merrylight Enterprises during this period, they shall be remunerated by Merrylight Enterprises in accordance with the applicable National Employment Council (NEC) for the Horticulture Industry wage rates and grades applicable to the duties performed. Payment shall be made only for the period worked for Merrylight Enterprises. At the time the driver resumes transporting employees from Merrylight Enterprises to their designated drop-off points at the end of the working day, the driver shall once again be regarded as performing duties under Jamport until the completion of the final employee drop-off. Nothing in this clause shall be construed as creating continuous employment with Merrylight Enterprises for the entire period the driver is present on the premises. Employment responsibilities and remuneration shall apply only for the periods during which work is voluntarily performed for Merrylight Enterprises. The voluntary work performed for Merrylight Enterprises shall be recorded on the appropriate attendance records and paid separately from the transport services provided by Jamport.
          </div>
        )}
      </div>

      {/* SIGNATURES */}
      <div style={{ display: "flex", gap: 24, flexWrap: "wrap", marginBottom: 16 }}>
        <SigBlock label="Employee" sigKey="employeeSig" nameKey="employeeSigName" dateKey="employeeSigDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
        <SigBlock label="Employer" sigKey="employerSig" nameKey="employerSigName" dateKey="employerSigDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
      </div>
      <SigBlock label="Witness" sigKey="witnessSig" nameKey="witnessName" dateKey="witnessDate" form={form} onChange={onChange} readOnly={readOnly} canEdit={canEditFields} />
      <p style={{fontSize: 11, color: "#666", marginTop: 12}}>NOTE: The witness should be conversant in English and the mother language of the employee and should confirm that the employee understands the content.</p>

      {!readOnly && <div style={{ display: "flex", gap: 8, marginTop: 14 }}><button className="btn btn-primary" style={{ width: "auto" }} onClick={() => onSave(form)}>Submit &amp; save</button><button className="btn btn-mini" onClick={onCancel}>Cancel</button></div>}
    </div>
    </EditContractCtx.Provider>
  );
}

function TruckDriverContractForm(props) { return <SeasonalContractForm {...props} isTruckDriver={true} />; }

/* ========== REGISTER FORMS INTO HR_FORM_TYPES ========== */
const hrPdfName = (prefix, f) => {
  const name = (f.employeeName || (f.surname ? (f.surname + (f.firstName ? "_" + f.firstName : "")) : "")).trim().replace(/\s+/g,"_") || "Unknown";
  return prefix + "_" + name;
};
HR_FORM_TYPES["contract"]     = { label: "Contract of Employment",                   component: ContractForm,          new: newContractForm,          pdfName: (f) => hrPdfName("Contract", f) };
HR_FORM_TYPES["permcontract"] = { label: "Permanent Contract of Employment",          component: PermanentContractForm,  new: newPermanentContractForm,  pdfName: (f) => hrPdfName("Permanent_Contract", f) };
HR_FORM_TYPES["seasonal"]     = { label: "Seasonal Contract of Employment",           component: SeasonalContractForm,   new: newSeasonalContractForm,   pdfName: (f) => hrPdfName("Seasonal_Contract", f) };
HR_FORM_TYPES["truckdriver"]  = { label: "Seasonal Contract — Truck Drivers",         component: TruckDriverContractForm, new: () => ({ ...newSeasonalContractForm(), id: "truckdriver-" + Date.now(), type: "truckdriver", version: CONTRACT_VERSIONS.truckdriver, customClauses: {}, jamportClause: "SPECIAL CLAUSE 19: HOURS OF WORK - TRANSPORT DRIVERS (JAMPORT): Employees engaged as transport drivers through Jamport shall be regarded as employees of Jamport while transporting employees from designated pick-up points to Merrylight Enterprises at commencement of working day. Upon arrival at Merrylight Enterprises and after parking their vehicles, the drivers shall remain at the workplace until the scheduled employee collection time. During this period, the drivers may, at their own discretion, elect to perform driving or other assigned duties for Merrylight Enterprises. Such work shall be entirely voluntary. No driver shall be compelled, coerced, or disadvantaged for choosing not to undertake such duties. Where a driver voluntarily agrees to perform duties for Merrylight Enterprises during this period, they shall be remunerated by Merrylight Enterprises in accordance with the applicable National Employment Council (NEC) for the Horticulture Industry wage rates and grades applicable to the duties performed. Payment shall be made only for the period worked for Merrylight Enterprises. At the time the driver resumes transporting employees from Merrylight Enterprises to their designated drop-off points at the end of the working day, the driver shall once again be regarded as performing duties under Jamport until the completion of the final employee drop-off. Nothing in this clause shall be construed as creating continuous employment with Merrylight Enterprises for the entire period the driver is present on the premises. Employment responsibilities and remuneration shall apply only for the periods during which work is voluntarily performed for Merrylight Enterprises. The voluntary work performed for Merrylight Enterprises shall be recorded on the appropriate attendance records and paid separately from the transport services provided by Jamport." }), pdfName: (f) => hrPdfName("TruckDriver_Contract", f) };

/* ================= APP SHELL ================= */
const NAV = [
  ["dash", "Dashboard", "◧", null],
  ["batches", "Batches", "≣", null],
  ["new", "Batch upload", "＋", ["admin", "entry"]],
  ["quality", "Quality Defects", "◬", null],
  ["season", "Season", "◔", null],
  ["packshed", "Packshed", "▤", null],
  ["receiving", "Fruit Receiving", "⬇", ["admin", "entry"]],
  ["packing", "Production Packing", "📦", ["admin", "entry"]],
  ["master", "Master Files", "⚙", ["admin"]],
  ["chat", "Chat", "◈", null],
  ["compliance", "Compliance", "✓", ["admin","qc","entry"]],
  ["hr", "Human Resources", "▤", ["admin","hrOfficer","hrManager"]],
  ["users", "Users", "◉", ["admin"]],
];

export default function App() {
  const seed = useMemo(() => SEED.rows.map(rowToRec), []);
  const [userRecs, setUserRecs] = useState(null);
  const [users, setUsers] = useState(null);
  const [me, setMe] = useState(undefined); // undefined = loading, null = signed out
  const [firstRun, setFirstRun] = useState(false);
  const [page, setPage] = useState("dash");
  const [navOpen, setNavOpen] = useState(false); // portrait-phone drawer
  const [updateAvail, setUpdateAvail] = useState(false);
  // Poll sw.js every 5 min to detect a new deployment
  useEffect(() => {
    const check = async () => {
      try {
        const res = await fetch("/sw.js?" + Date.now(), { cache: "no-store" });
        const text = await res.text();
        const m = text.match(/mlerp-v(\d+)/);
        if (m && parseInt(m[1]) > 81) setUpdateAvail(true);
      } catch (e) {}
    };
    check();
    const t = setInterval(check, 5 * 60 * 1000);
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    (async () => {
      let list = await loadUsersStore();
      if (!list || !list.length) {
        list = [{ name: "Administrator", username: "admin", role: "admin", pin: await sha("1234"), created: Date.now() }];
        await saveUsersStore(list);
        setFirstRun(true);
      }
      setUsers(list);
      // Always open on the login screen: sessions are never restored.
      // Shared packhouse devices must not wake up already signed in.
      clearSession();
      setMe(null);
      setUserRecs(await loadUser());
    })();
  }, []);

  // Live refresh: pick up other devices' changes without restarting the app.
  useEffect(() => {
    const typing = () => {
      const el = document.activeElement;
      return !!el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.tagName === "SELECT");
    };
    const refresh = async () => {
      if (typing()) return; // never yank state while someone is mid-entry
      try {
        const [freshRecs, freshUsers] = await Promise.all([loadUser(), loadUsersStore()]);
        if (freshRecs) setUserRecs((cur) => (JSON.stringify(freshRecs) !== JSON.stringify(cur) ? freshRecs : cur));
        if (freshUsers && freshUsers.length) {
          setUsers((cur) => (JSON.stringify(freshUsers) !== JSON.stringify(cur) ? freshUsers : cur));
          let removed = false;
          setMe((cur) => {
            if (!cur) return cur;
            const mine = freshUsers.find((x) => x.username === cur.username);
            if (!mine) { removed = true; return null; }
            return JSON.stringify(mine) !== JSON.stringify(cur) ? mine : cur;
          });
          if (removed) clearSession();
        }
      } catch (e) { /* offline or transient — next tick will retry */ }
    };
    const t = setInterval(refresh, 20000);
    const onWake = () => refresh();
    window.addEventListener("focus", onWake);
    document.addEventListener("visibilitychange", onWake);
    window.addEventListener("online", onWake);
    return () => {
      clearInterval(t);
      window.removeEventListener("focus", onWake);
      document.removeEventListener("visibilitychange", onWake);
      window.removeEventListener("online", onWake);
    };
  }, []);

  // Drawer swipe gestures (portrait phones): from the left edge opens; swipe left closes.
  useEffect(() => {
    let sx = null, sy = null;
    const ts = (e) => { const t = e.touches[0]; sx = t.clientX; sy = t.clientY; };
    const te = (e) => {
      if (sx == null) return;
      const t = e.changedTouches[0];
      const dx = t.clientX - sx, dy = Math.abs(t.clientY - sy);
      if (dy < 60) {
        if (sx < 30 && dx > 60) setNavOpen(true);
        else if (dx < -60) setNavOpen(false);
      }
      sx = sy = null;
    };
    document.addEventListener("touchstart", ts, { passive: true });
    document.addEventListener("touchend", te, { passive: true });
    return () => { document.removeEventListener("touchstart", ts); document.removeEventListener("touchend", te); };
  }, []);

  const recs = useMemo(() => (userRecs ? [...seed, ...userRecs] : seed), [seed, userRecs]);

  const login = async (username, pin) => {
    const acc = (users || []).find((x) => x.username === username.trim().toLowerCase());
    if (!acc) return "No user with that username.";
    if (acc.pin !== (await sha(pin))) return "Incorrect PIN.";
    setMe(acc); // in-memory only — closing the app signs you out
    return null;
  };
  const logout = () => { setMe(null); setPage("dash"); clearSession(); };

  const resetAccess = async () => {
    const list = [{ name: "Administrator", username: "admin", role: "admin", pin: await sha("1234"), created: Date.now() }];
    setUsers(list);
    await saveUsersStore(list);
    setFirstRun(true);
  };

  const updateUsers = async (list) => {
    setUsers(list);
    const ok = await saveUsersStore(list);
    const mine = list.find((x) => x.username === me.username);
    if (!mine) logout(); else setMe(mine);
    return ok;
  };

  const deleteAppBatches = async () => {
    if (!me || me.role !== "admin") return false; // deleting uploads is admin-only
    setUserRecs([]);
    return await saveUser([]);
  };
  const addRecs = async (list) => {
    const stamp = Date.now();
    const withIds = list.map((rec, i) => ({ ...rec, id: "u-" + stamp + "-" + i, source: "app", enteredBy: me ? me.username : null }));
    const next = [...(userRecs || []), ...withIds];
    setUserRecs(next);
    return await saveUser(next);
  };
  const delRec = async (id) => {
    if (!me || me.role !== "admin") return false; // deleting uploads is admin-only
    const next = (userRecs || []).filter((r) => r.id !== id);
    setUserRecs(next);
    await saveUser(next);
  };

  if (me === undefined) {
    return (
      <div className="app"><style>{CSS}</style>
        <div className="login-wrap"><p className="muted">Loading…</p></div>
      </div>
    );
  }
  if (!me) {
    return (
      <div className="app"><style>{CSS}</style>
        <Login onLogin={login} onReset={resetAccess} firstRun={firstRun} />
      </div>
    );
  }

  const allowed = (roles) => !roles || roles.includes(me.role);
  const pageOk = (k) => me.role === "admin" || !me.pages || me.pages.includes(k);
  let nav = NAV.filter(([k, , , roles]) => allowed(roles) && pageOk(k));
  if (!nav.length) nav = NAV.filter(([, , , roles]) => allowed(roles)); // never strand a user with zero pages
  const safePage = nav.some(([k]) => k === page) ? page : (nav.length ? nav[0][0] : "dash");

  return (
    <div className={"app" + (navOpen ? " nav-open" : "")}>
      <style>{CSS}</style>
      <button className="burger" aria-label="Menu" onClick={() => setNavOpen(!navOpen)}>☰</button>
      {navOpen && <div className="nav-veil" onClick={() => setNavOpen(false)} />}
      <nav className="side">
        <div className="brand">
          <img className="brand-logo" src={LOGO} alt="Merrylight Enterprises" />
          <div>
            <div className="brand-name">MERRYLIGHT</div>
            <div className="brand-sub">Packhouse ERP · Season 2026</div>
          </div>
        </div>
        <div className="nav">
          {nav.map(([k, l, ic]) => (
            <button key={k} className={"nav-btn" + (safePage === k ? " active" : "")} onClick={() => { setPage(k); setNavOpen(false); }}>
              <span className="nav-ic">{ic}</span>{l}
            </button>
          ))}
        </div>
        <div className="side-foot">
          <div className="who">
            <div className="who-name">{me.name}</div>
            <div className="who-role">{ROLE_LABEL[me.role]}</div>
          </div>
          <button className="btn btn-ghost" onClick={logout}>Sign out</button>
          <div className="app-ver">v81 · 17 Jul</div>
        </div>
        {updateAvail && (
          <div className="update-banner" onClick={() => window.location.reload(true)}>
            🔄 Update available — tap to refresh
          </div>
        )}
      </nav>

      <main className="main">
        <div className="print-head">
          <img src={LOGO} alt="" style={{ width: 34, height: 34, borderRadius: "50%" }} />
          <div>
            <div style={{ fontFamily: "'Bricolage Grotesque'", fontWeight: 800, fontSize: 15 }}>MERRYLIGHT PACKHOUSE ERP</div>
            <div style={{ fontSize: 10.5, color: "#6E6D85" }}>Printed by <b>{me.name}</b> ({ROLE_LABEL[me.role]}) · {new Date().toLocaleString("en-GB", { day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })}</div>
          </div>
        </div>
        {safePage === "dash" && <Dashboard recs={recs} />}
        {safePage === "batches" && <Batches recs={recs} onDelete={delRec} me={me} />}
        {safePage === "new" && <NewBatch recs={recs} onSaveMany={addRecs} onDeleteApp={deleteAppBatches} me={me} />}
        {safePage === "quality" && <Quality recs={recs} />}
        {safePage === "season" && <Season recs={recs} me={me} />}
        {safePage === "packshed" && <Packshed me={me} />}
        {safePage === "receiving" && <FruitReceiving me={me} />}
        {safePage === "packing" && <ProductionPacking me={me} />}
        {safePage === "master" && <MasterFiles me={me} />}
        {safePage === "chat" && <Chat me={me} />}
        {safePage === "compliance" && <ComplianceErrorBoundary><Compliance me={me} /></ComplianceErrorBoundary>}
        {safePage === "hr" && <HRErrorBoundary><HumanResources me={me} /></HRErrorBoundary>}
        {safePage === "users" && <Users me={me} users={users || []} onChange={updateUsers} />}
      </main>
    </div>
  );
}

/* ================= STYLES ================= */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:wght@500;700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500&display=swap');
*{box-sizing:border-box;margin:0}
:root{
  --ink:#232047;--muted:#6E6D85;--paper:#F3F4F9;--surface:#fff;--line:#E4E5F0;
  --bloom:#4C46A8;--deep:#201C40;--leaf:#2F7D53;--amber:#B4761F;--red:#AE4038;
}
.app{display:flex;min-height:100vh;background:var(--paper);color:var(--ink);
  font-family:'Inter',system-ui,sans-serif;font-size:13.5px;line-height:1.45}
h2{font-family:'Bricolage Grotesque',sans-serif;font-weight:800;font-size:22px;letter-spacing:-.01em}
h3{font-family:'Bricolage Grotesque',sans-serif;font-weight:700;font-size:13px;letter-spacing:.02em;text-transform:uppercase;color:var(--ink)}
h4{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:6px}
.muted{color:var(--muted)} .small{font-size:11px} .center{text-align:center;padding:24px 0!important}
.mono{font-family:'IBM Plex Mono',monospace;font-size:12px}
.strong{font-weight:600} .r{text-align:right} .pos{color:var(--leaf)} .neg{color:var(--red)}

/* sidebar */
.side{width:212px;background:var(--deep);color:#CBCBE4;display:flex;flex-direction:column;padding:20px 12px;position:sticky;top:0;height:100vh;flex-shrink:0}
.brand{display:flex;gap:10px;align-items:center;padding:0 8px 18px;border-bottom:1px solid rgba(255,255,255,.09);margin-bottom:14px}
.brand-logo{width:38px;height:38px;border-radius:50%;object-fit:cover;flex:none}
.login-brand .brand-logo{width:52px;height:52px}
.brand-mark{display:grid;grid-template-columns:8px 8px;gap:3px}
.brand-mark span{width:8px;height:8px;border-radius:50%}
.brand-mark span:nth-child(1){background:#B9BEEC}.brand-mark span:nth-child(2){background:#7F86DE}
.brand-mark span:nth-child(3){background:#7F86DE}.brand-mark span:nth-child(4){background:#4C46A8}
.brand-name{font-family:'Bricolage Grotesque';font-weight:800;color:#fff;font-size:15px;letter-spacing:.06em}
.brand-sub{font-size:9.5px;letter-spacing:.05em;color:#8D8CB0;text-transform:uppercase}
.nav{display:flex;flex-direction:column;gap:2px}
.nav-btn{display:flex;align-items:center;gap:10px;width:100%;text-align:left;background:none;border:0;color:#B9B8D4;
  padding:9px 10px;border-radius:8px;font:inherit;font-weight:500;cursor:pointer}
.nav-btn:hover{background:rgba(255,255,255,.06);color:#fff}
.nav-btn.active{background:var(--bloom);color:#fff;font-weight:600}
.nav-ic{width:16px;text-align:center;opacity:.85}
.side-foot{margin-top:auto;font-size:10.5px;color:#7C7BA0;padding:10px 8px 0;border-top:1px solid rgba(255,255,255,.09)}
.app-ver{font-size:9px;color:#6A697E;margin-top:6px;letter-spacing:.04em}
.update-banner{position:fixed;bottom:0;left:0;right:0;background:#4C46A8;color:#fff;text-align:center;padding:10px 16px;font-size:13px;font-weight:600;cursor:pointer;z-index:200;animation:slide-up .3s ease}
@keyframes slide-up{from{transform:translateY(100%)}to{transform:none}}

/* main */
.main{flex:1;padding:22px 26px 50px;width:100%;min-width:0;max-width:1280px;margin:0 auto}
.page-head{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap;margin-bottom:14px}
.page-head p{margin-top:3px}
.controls{display:flex;gap:8px;align-items:center}
.controls.wrap{flex-wrap:wrap}
.lbl{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted)}
.sel,.inp{border:1px solid var(--line);background:var(--surface);border-radius:8px;padding:7px 10px;font:inherit;color:var(--ink)}
.sel:focus,.inp:focus,.btn:focus{outline:2px solid var(--bloom);outline-offset:1px}

/* cards & kpis */
.card{background:var(--surface);border:1px solid var(--line);border-radius:12px;margin-bottom:13px;overflow:hidden}
.card-head{display:flex;justify-content:space-between;align-items:center;padding:9px 14px;border-bottom:1px solid var(--line)}
.card-body{padding:16px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:9px;margin-bottom:13px}
.kpi{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:9px 12px}
.kpi-label{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
.kpi-value{font-family:'Bricolage Grotesque';font-weight:800;font-size:21px;margin-top:3px;letter-spacing:-.01em}
.kpi-sub{font-size:11px;color:var(--muted);margin-top:2px}
.kpi-good .kpi-value{color:var(--leaf)} .kpi-bad .kpi-value{color:var(--red)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:900px){.grid2{grid-template-columns:1fr}}

/* tables */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:600;
  text-align:left;padding:5px 10px;border-bottom:1px solid var(--line);white-space:nowrap;background:#FAFAFE}
th.r{text-align:right}
td{padding:5px 10px;border-bottom:1px solid #EFF0F7;white-space:nowrap;vertical-align:middle;font-size:12.5px;line-height:1.35}
tbody tr:last-child td{border-bottom:0}
.rowbtn{cursor:pointer} .rowbtn:hover td{background:#F6F6FC}
.totalrow td{background:#EFEFF8;border-top:2px solid var(--bloom)}
.dayhead td{background:var(--bloom);color:#fff;font-weight:700;font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;padding:5px 10px}
.hidef td{background:#FCF0EF;box-shadow:inset 0 -1px 0 #F3DBD8}
.hidef td:first-child{box-shadow:inset 3px 0 0 var(--red),inset 0 -1px 0 #F3DBD8}
.rowbtn.hidef:hover td{background:#F8E4E2}
.userrow.hidef td:first-child{box-shadow:inset 3px 0 0 var(--bloom),inset 0 -1px 0 #F3DBD8}
.userrow td:first-child{box-shadow:inset 3px 0 0 var(--bloom)}
.badge{display:inline-block;margin-left:6px;font-size:9px;font-weight:700;letter-spacing:.05em;color:#fff;
  background:var(--bloom);border-radius:4px;padding:1.5px 5px;vertical-align:1px}
.chev{color:var(--muted);width:24px}
.detail td{background:#F8F8FD;white-space:normal}
.detail-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:18px;padding:6px 2px}
.detail-grid dl{display:grid;grid-template-columns:auto 1fr;gap:3px 12px;font-size:12px}
.detail-grid dt{color:var(--muted)} .detail-grid dd{text-align:right;font-weight:500}
.detail-actions{display:flex;align-items:flex-end;gap:8px;flex-wrap:wrap}
.batch-edit{margin-top:14px;padding:14px;background:#F6F6FD;border-radius:10px;border:1px solid var(--line)}

.dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px;vertical-align:-.5px}

/* entry form */
.entry-grid{display:grid;grid-template-columns:1fr 280px;gap:16px;align-items:start}
@media(max-width:900px){.entry-grid{grid-template-columns:1fr}}
.fgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}
.fgrid4{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:10px}
.fld{display:flex;flex-direction:column;gap:4px}
.fld>span{font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted)}
.fld .inp{width:100%}
.calc{display:grid;grid-template-columns:1fr auto;gap:7px 10px;font-size:12.5px}
.calc dt{color:var(--muted)} .calc dd{text-align:right;font-weight:600}
.checks{list-style:none;margin-bottom:12px;font-size:12.5px}
.checks li{padding:3px 0}
.checks .ok{color:var(--leaf)} .checks .todo{color:var(--muted)} .checks .warn{color:var(--amber);font-weight:600}
.btn{border:0;border-radius:8px;padding:9px 14px;font:inherit;font-weight:600;cursor:pointer}
.btn-primary{background:var(--bloom);color:#fff;width:100%}
.btn-primary:disabled{background:#C6C5DE;cursor:not-allowed}
.btn-danger{background:#FBEDEC;color:var(--red)}
.savedmsg{margin-top:10px;font-size:12px}
.savedmsg.ok{color:var(--leaf)} .savedmsg.warn{color:var(--amber)}

/* login & auth */
.login-wrap{flex:1;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}
.login-card{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:26px 24px;width:340px;max-width:100%;display:flex;flex-direction:column;gap:12px;box-shadow:0 10px 34px rgba(35,32,71,.08)}
.login-brand{border:0;padding:0;margin:0 0 4px}
.login-err{color:var(--red);font-size:12.5px;font-weight:600}
.login-hint{font-size:12px;color:var(--amber);background:#FBF4E7;border-radius:8px;padding:8px 10px}
.login-note{font-size:11px;color:var(--muted);line-height:1.5}
.who{margin-bottom:8px}
.who-name{color:#fff;font-weight:600;font-size:12.5px}
.who-role{font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:#8D8CB0}
.btn-ghost{background:rgba(255,255,255,.08);color:#CBCBE4;width:100%;padding:7px}
.btn-ghost:hover{background:rgba(255,255,255,.14);color:#fff}
.btn-mini{padding:4px 9px;font-size:11px;background:#EFEFF8;color:var(--ink)}
.btn-mini.btn-danger{background:#FBEDEC;color:var(--red)}
.btn-mini:disabled{opacity:.4;cursor:not-allowed}

/* team chat */
.compliance-form table{width:100%;border-collapse:collapse}
.badge.ok{background:#D4EDDA;color:#1A5C2A;padding:2px 7px;border-radius:8px;font-size:11px}
.badge.err{background:#F8D7DA;color:#842029;padding:2px 7px;border-radius:8px;font-size:11px}
.chat{display:flex;flex-direction:column;height:calc(100vh - 190px);min-height:380px}
.chat-msgs{flex:1;overflow-y:auto;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px;display:flex;flex-direction:column;gap:12px}
.msg{max-width:640px}
.msg.mine{align-self:flex-end;text-align:right}
.msg-meta{font-size:10.5px;color:var(--muted);margin-bottom:3px;display:flex;gap:8px;align-items:center}
.msg.mine .msg-meta{justify-content:flex-end}
.msg-meta b{color:var(--ink)}
.msg-role{text-transform:uppercase;letter-spacing:.04em;font-size:9px;background:#EFEFF8;border-radius:4px;padding:1px 5px}
.msg-bubble{display:inline-block;background:#F1F1F9;border-radius:12px;padding:9px 13px;font-size:13px;line-height:1.45;text-align:left;white-space:pre-wrap;word-break:break-word}
.msg.mine .msg-bubble{background:var(--bloom);color:#fff;border-bottom-right-radius:4px}
.msg:not(.mine) .msg-bubble{border-bottom-left-radius:4px}
.msg-del{background:none;border:0;color:var(--muted);cursor:pointer;font-size:13px;padding:0 2px;line-height:1}
.msg-del:hover{color:var(--red)}
.chat-input{display:flex;gap:8px;margin-top:12px}

/* file import */
.bw-in{width:76px;padding:4px 7px;font-size:12px;text-align:right}
.filein{font:inherit;font-size:12.5px;color:var(--ink)}
.filein::file-selector-button{border:0;border-radius:8px;padding:8px 13px;margin-right:10px;font:inherit;font-weight:600;background:var(--bloom);color:#fff;cursor:pointer}

/* sub tabs */
.tabs{display:flex;gap:2px;border-bottom:1px solid var(--line);margin-bottom:16px}
.tab{background:none;border:0;font:inherit;font-weight:600;font-size:13px;color:var(--muted);
  padding:9px 14px;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
.tab:hover{color:var(--ink)}
.tab.active{color:var(--bloom);border-bottom-color:var(--bloom)}

/* login help */
.linklike{background:none;border:0;color:var(--bloom);font:inherit;font-size:12px;cursor:pointer;text-decoration:underline;padding:0;text-align:left}
.login-help{background:#F6F6FB;border:1px solid var(--line);border-radius:8px;padding:10px 12px;display:flex;flex-direction:column;gap:8px;font-size:12px;color:var(--muted);line-height:1.5}

/* order editor */
.oed{padding:14px 16px;background:#F8F8FD;border-bottom:1px solid var(--line);white-space:normal}
.oed-days{display:grid;grid-template-columns:1fr 1fr auto;gap:18px;margin-top:12px;align-items:end}
@media(max-width:760px){.oed-days{grid-template-columns:1fr}}
.fgrid3{display:grid;grid-template-columns:repeat(3,minmax(70px,1fr));gap:8px}
.oed-bal{min-width:110px}
.oed-balv{font-family:'Bricolage Grotesque';font-weight:800;font-size:18px}
.oed-actions{display:flex;gap:8px;align-items:center;margin-top:14px}
.oed-kg{margin-top:6px;font-size:11.5px;color:var(--muted);font-family:'IBM Plex Mono',monospace}

/* chart tooltip */
.tip{background:#fff;border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:11.5px;box-shadow:0 4px 14px rgba(35,32,71,.08)}
.tip-t{font-weight:700;margin-bottom:4px}
.tip-r{display:flex;align-items:center;gap:2px;padding:1px 0}

/* mobile */
.burger{display:none}
.nav-veil{display:none}
/* portrait phones: sidebar becomes a swipe-in drawer over the content */
@media(max-width:760px) and (orientation:portrait){
  .burger{display:flex;align-items:center;justify-content:center;position:fixed;top:10px;left:10px;z-index:60;
    width:40px;height:40px;border:0;border-radius:10px;background:var(--deep);color:#fff;font-size:17px;cursor:pointer;
    box-shadow:0 4px 14px rgba(32,28,64,.35)}
  .side{position:fixed;top:0;left:0;bottom:0;height:100vh;width:242px;z-index:55;
    transform:translateX(-105%);transition:transform .22s ease;box-shadow:6px 0 24px rgba(32,28,64,.3)}
  .nav-open .side{transform:none}
  .nav-veil{display:block;position:fixed;inset:0;background:rgba(32,28,64,.45);z-index:50}
  .main{padding:58px 14px 50px}
}
/* landscape phones: sidebar stays put as a compact rail */
@media(max-width:960px) and (orientation:landscape) and (max-height:520px){
  .side{width:172px;padding:12px 8px}
  .brand-sub{display:none}
  .nav-btn{padding:7px 9px;font-size:12.5px}
  .side-foot{display:none}
  .main{padding:14px 14px 40px}
}
@media(prefers-reduced-motion:no-preference){
  .nav-btn,.rowbtn td{transition:background .12s ease}
}

/* snapshot mode (direct PDF download) */
.snap .print-head{display:flex!important;align-items:center;gap:10px;margin-bottom:12px;border-bottom:2px solid #4C46A8;padding-bottom:8px}
.snap .controls,.snap .tabs,.snap .btn,.snap .filein,.snap .savedmsg,.snap .chat-input,.snap .linklike{display:none!important}
.snap .main{max-width:none;width:max-content;min-width:1000px}
.snap .tbl-wrap{overflow:visible!important}
.snap .kpis{grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:7px;margin-bottom:10px}
.snap .card{margin-bottom:10px}
.snap td{padding:3px 8px;font-size:11.5px}
.snap th{padding:3px 8px}

/* print / PDF export */
.print-head{display:none}
@media print{
  .side,.burger,.nav-veil,.controls,.tabs,.btn,.filein,.chat-input,.oed,.savedmsg,.linklike{display:none!important}
  .app{display:block;background:#fff}
  .main{max-width:none;margin:0;padding:0}
  .print-head{display:flex!important;align-items:center;gap:10px;margin-bottom:14px;border-bottom:2px solid #4C46A8;padding-bottom:8px}
  .card,.kpi{box-shadow:none;break-inside:avoid;border-color:#D8D9E6}
  .kpis{break-inside:avoid}
  .grid2{break-inside:avoid}
  .tbl-wrap{overflow:visible}
  table{font-size:10px}
  th,td{padding:3px 6px}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .page-head h2{font-size:18px}
}
`;
