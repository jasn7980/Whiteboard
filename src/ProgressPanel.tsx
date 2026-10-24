import { useEffect, useState } from 'react'

// ---------------------------------------------------------------------------
// Panel "Progreso": la tarjeta viva del seguimiento de Jaime.
// Lee public/progreso.json (generado desde sus CSV de TrainingPeaks) y
// muestra: la tendencia de sus maratones 2026 hacia la meta 3:45, el
// volumen semanal, la carrera objetivo (Ottawa o Chicago 2027, a elegir)
// y sus sesiones de intervalos con el detalle completo de cada una.
// ---------------------------------------------------------------------------

type Race = { id: string; name: string; date: string; note: string }
type ProgressData = {
  updated: string
  goal: { time: string; pace: string; races: Race[] }
  marathons: { race: string; date: string; time: string; seconds: number; pace: string }[]
  half: { race: string; date: string; time: string }
  personalBests: { race: string; date: string; time: string; note: string }[]
  weekly: { week: string; km: number; tss: number }[]
  intervalSessions: {
    fecha: string; bloque: string; nombre: string; km: string;
    duracion: string; ritmo: string; tss: string; instrucciones: string
  }[]
  progress: { from: string; now: string; goal: string; gained: string; remaining: string; pct: number }
  totals: { workouts: number; ottawa_km: number; petit_km: number }
}

function daysUntil(iso: string): number {
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  const target = new Date(`${iso}T00:00:00`)
  return Math.round((target.getTime() - today.getTime()) / 86400000)
}

function TrendChart({ data }: { data: ProgressData }) {
  const W = 320
  const H = 132
  const pts = [
    ...data.marathons.map((m) => ({ label: m.race, time: m.time, v: m.seconds / 60 })),
    { label: 'Meta', time: data.goal.time, v: 225 },
  ]
  const min = 222
  const max = 246
  const x = (i: number) => 34 + (i * (W - 50)) / (pts.length - 1)
  const y = (v: number) => 8 + ((max - v) / (max - min)) * (H - 30)
  const line = pts.map((p, i) => `${i === 0 ? 'M' : 'L'}${x(i)},${y(p.v)}`).join(' ')
  return (
    <svg width="100%" viewBox={`0 0 ${W} ${H}`} style={{ display: 'block' }}>
      <line x1={34} y1={y(225)} x2={W - 8} y2={y(225)} stroke="#c62828" strokeDasharray="4 3" strokeWidth={1} />
      <text x={W - 8} y={y(225) - 4} fontSize={9} fill="#c62828" textAnchor="end">meta 3:45</text>
      <path d={line} fill="none" stroke="#1565c0" strokeWidth={2.5} strokeLinejoin="round" />
      {pts.map((p, i) => (
        <g key={p.label}>
          <circle cx={x(i)} cy={y(p.v)} r={4} fill={i === pts.length - 1 ? '#c62828' : '#1565c0'} />
          <text x={x(i)} y={y(p.v) - 8} fontSize={10} fontWeight={700} textAnchor="middle" fill="#111">
            {p.time}
          </text>
          <text x={x(i)} y={H - 4} fontSize={9} textAnchor="middle" fill="#555">
            {p.label}
          </text>
        </g>
      ))}
    </svg>
  )
}

function WeekBars({ data }: { data: ProgressData }) {
  const maxKm = Math.max(...data.weekly.map((w) => w.km), 1)
  return (
    <div style={{ display: 'flex', alignItems: 'flex-end', gap: 1.5, height: 54 }}>
      {data.weekly.map((w) => (
        <div
          key={w.week}
          title={`${w.week}: ${w.km} km · ${w.tss} TSS`}
          style={{
            flex: 1,
            height: `${Math.max(4, (w.km / maxKm) * 54)}px`,
            background: '#66bb6a',
            borderRadius: '2px 2px 0 0',
          }}
        />
      ))}
    </div>
  )
}

export default function ProgressPanel() {
  const [data, setData] = useState<ProgressData | null>(null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    fetch('progreso.json')
      .then((r) => {
        if (!r.ok) throw new Error(`progreso.json: ${r.status}`)
        return r.json()
      })
      .then(setData)
      .catch((e) => setError(e instanceof Error ? e.message : String(e)))
  }, [])

  if (error) {
    return (
      <div style={{ fontSize: 13, color: '#7f0000', padding: 8 }}>
        No pude cargar tu progreso: {error}
      </div>
    )
  }
  if (!data) {
    return <div style={{ fontSize: 13, color: '#555', padding: 8 }}>Cargando tu progreso…</div>
  }

  const recent = [...data.intervalSessions].reverse().slice(0, 12)
  const ottawaDays = daysUntil(data.goal.races[0].date)
  const chicagoDays = daysUntil(data.goal.races[1].date)
  const verdunDays = daysUntil('2027-09-12')

  return (
    <div
      style={{
        fontSize: 13,
        color: '#222',
        display: 'flex',
        flexDirection: 'column',
        gap: 10,
        flex: 1,
        minHeight: 0,
      }}
    >
      <div>
        <div style={{ fontSize: 15, fontWeight: 700 }}>🏃 Meta maratón: 3:45</div>
        <div style={{ color: '#555' }}>
          {data.progress.from} → {data.progress.now} · faltan {data.progress.remaining} ·{' '}
          {data.progress.pct}% del camino
        </div>
        <div style={{ background: '#e0e0e0', borderRadius: 6, height: 12, marginTop: 6, overflow: 'hidden' }}>
          <div style={{ width: `${data.progress.pct}%`, background: '#2e7d32', height: '100%' }} />
        </div>
        <div style={{ color: '#555', marginTop: 4 }}>
          Ritmo: 5:41 → 5:34 /km · meta 5:20 /km · Medio maratón: {data.half.time} (Montreal, abr 2026)
        </div>
        <div style={{ color: '#333', marginTop: 4 }}>
          ⭐ Tu mejor maratón: Lala 2019, 3:48 (Lala 2020: 3:52) · la meta 3:45 está a 3 minutos de tu marca
        </div>
        <div style={{ color: '#333', marginTop: 2 }}>
          ⭐ Tu mejor medio: 1:36:30 (ESPN, 5 may 2019)
        </div>
      </div>

      <div>
        <div style={{ fontWeight: 700, marginBottom: 2 }}>Maratones 2026</div>
        <TrendChart data={data} />
      </div>

      <div>
        <div style={{ fontWeight: 700, marginBottom: 2 }}>
          Km por semana (40 semanas · {data.totals.ottawa_km + data.totals.petit_km} km en 2026)
        </div>
        <WeekBars data={data} />
      </div>

      <div>
        <div style={{ fontWeight: 700, marginBottom: 4 }}>Plan B: en dos pasos hacia el 3:45</div>
        <div style={{ color: '#333' }}>Desde tu 3:54:51 de Petit Train:</div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4, marginTop: 6 }}>
          <div style={{ border: '1px solid #ddd', borderRadius: 8, padding: '6px 8px', background: '#fff' }}>
            <div style={{ fontWeight: 700 }}>
              1 · {data.goal.races[0].name} — {data.goal.races[0].date} · faltan{' '}
              {ottawaDays} días · <strong>{Math.floor(ottawaDays / 7)} semanas</strong>
            </div>
            <div style={{ color: '#555' }}>
              Objetivo ~3:50 (bajar 5 minutos) · {data.goal.races[0].note}
            </div>
          </div>
          <div style={{ textAlign: 'center', color: '#888' }}>↓</div>
          <div
            style={{
              border: '2px solid #1565c0',
              borderRadius: 8,
              padding: '6px 8px',
              background: '#e3f2fd',
            }}
          >
            <div style={{ fontWeight: 700 }}>
              2 · {data.goal.races[1].name} — {data.goal.races[1].date} · faltan{' '}
              {chicagoDays} días · <strong>{Math.floor(chicagoDays / 7)} semanas</strong>
            </div>
            <div style={{ color: '#333' }}>
              Objetivo 3:45 (otros 5 minutos) · carrera plana, ideal para el remate
            </div>
            <div style={{ color: '#555', fontSize: 12 }}>{data.goal.races[1].note}</div>
          </div>
        </div>
        <div
          style={{
            marginTop: 6,
            fontSize: 12,
            color: '#555',
            borderTop: '1px dashed #ccc',
            paddingTop: 4,
          }}
        >
          Si no sales en la rifa de Chicago (resultados 8 dic 2026), el plan pasa a:{' '}
          <strong>Maratón BougeBouge Verdun — 12 sep 2027</strong> · faltan {verdunDays} días ·{' '}
          {Math.floor(verdunDays / 7)} semanas · mismo objetivo 3:45
        </div>
      </div>

      <div
        style={{
          flex: '1 1 auto',
          minHeight: 0,
          overflowY: 'auto',
          touchAction: 'pan-y',
          paddingRight: 2,
        }}
      >
        <div style={{ fontWeight: 700, marginBottom: 4 }}>
          Intervalos 2026 ({data.intervalSessions.length} sesiones · detalle completo)
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          {recent.map((s) => (
            <details key={`${s.fecha}-${s.nombre}`} style={{ border: '1px solid #ddd', borderRadius: 8, padding: '4px 8px', background: '#fff' }}>
              <summary style={{ cursor: 'pointer', fontWeight: 600 }}>
                {s.fecha} · {s.nombre} <span style={{ fontWeight: 400, color: '#555' }}>({s.km} km · {s.duracion})</span>
              </summary>
              <div style={{ marginTop: 4, color: '#333' }}>{s.instrucciones}</div>
              <div style={{ color: '#555', fontSize: 12 }}>
                Ritmo prom {s.ritmo} · TSS {s.tss} · bloque {s.bloque}
              </div>
            </details>
          ))}
        </div>
        <div style={{ color: '#555', fontSize: 12, marginTop: 4 }}>
          Datos: tus archivos de TrainingPeaks ({data.totals.workouts} entrenos detallados) · actualizado {data.updated}
        </div>
      </div>
    </div>
  )
}
