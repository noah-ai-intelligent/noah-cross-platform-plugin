import { useMemo } from 'react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  LineChart,
  Line,
  AreaChart,
  Area,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend
} from 'recharts';

export interface ChartPreviewProps {
  jsonString: string;
}

const COLORS = ['#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];

export function ChartPreview({ jsonString }: ChartPreviewProps) {
  const chartData = useMemo(() => {
    try {
      const parsed = JSON.parse(jsonString);
      return parsed;
    } catch (e) {
      return null;
    }
  }, [jsonString]);

  if (!chartData || !chartData.type || !chartData.data) {
    return (
      <div className="p-3 text-xs text-red-500 bg-red-50 rounded-lg border border-red-100">
        Invalid chart data format.
      </div>
    );
  }

  const { type, title, data, xKey, series } = chartData;

  const renderChart = () => {
    switch (type) {
      case 'bar':
        return (
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={data} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
              <XAxis dataKey={xKey || 'name'} tick={{ fontSize: 10, fill: '#71717a' }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 10, fill: '#71717a' }} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{ borderRadius: '8px', border: '1px solid #e4e4e7', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                itemStyle={{ fontSize: '12px' }}
                labelStyle={{ fontSize: '12px', fontWeight: '600', color: '#18181b', marginBottom: '4px' }}
                cursor={{ fill: '#f4f4f5' }}
              />
              {series && series.length > 1 && <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />}
              {(series || []).map((s: any, i: number) => (
                <Bar key={s.key} dataKey={s.key} name={s.label || s.key} fill={COLORS[i % COLORS.length]} radius={[2, 2, 0, 0]} />
              ))}
            </BarChart>
          </ResponsiveContainer>
        );
      case 'line':
        return (
          <ResponsiveContainer width="100%" height={250}>
            <LineChart data={data} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
              <XAxis dataKey={xKey || 'name'} tick={{ fontSize: 10, fill: '#71717a' }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 10, fill: '#71717a' }} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{ borderRadius: '8px', border: '1px solid #e4e4e7', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                itemStyle={{ fontSize: '12px' }}
                labelStyle={{ fontSize: '12px', fontWeight: '600', color: '#18181b', marginBottom: '4px' }}
              />
              {series && series.length > 1 && <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />}
              {(series || []).map((s: any, i: number) => (
                <Line key={s.key} type="monotone" dataKey={s.key} name={s.label || s.key} stroke={COLORS[i % COLORS.length]} strokeWidth={2} dot={{ r: 3 }} activeDot={{ r: 5 }} />
              ))}
            </LineChart>
          </ResponsiveContainer>
        );
      case 'area':
        return (
          <ResponsiveContainer width="100%" height={250}>
            <AreaChart data={data} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
              <XAxis dataKey={xKey || 'name'} tick={{ fontSize: 10, fill: '#71717a' }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 10, fill: '#71717a' }} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{ borderRadius: '8px', border: '1px solid #e4e4e7', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                itemStyle={{ fontSize: '12px' }}
                labelStyle={{ fontSize: '12px', fontWeight: '600', color: '#18181b', marginBottom: '4px' }}
              />
              {series && series.length > 1 && <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />}
              {(series || []).map((s: any, i: number) => (
                <Area key={s.key} type="monotone" dataKey={s.key} name={s.label || s.key} fill={COLORS[i % COLORS.length]} stroke={COLORS[i % COLORS.length]} fillOpacity={0.2} />
              ))}
            </AreaChart>
          </ResponsiveContainer>
        );
      case 'pie':
      case 'donut':
        const valueKey = series && series.length > 0 ? series[0].key : 'value';
        return (
          <ResponsiveContainer width="100%" height={250}>
            <PieChart margin={{ top: 10, right: 10, left: 10, bottom: 10 }}>
              <Tooltip
                contentStyle={{ borderRadius: '8px', border: '1px solid #e4e4e7', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                itemStyle={{ fontSize: '12px' }}
              />
              <Legend wrapperStyle={{ fontSize: '11px' }} />
              <Pie
                data={data}
                nameKey={xKey || 'name'}
                dataKey={valueKey}
                cx="50%"
                cy="50%"
                innerRadius={type === 'donut' ? 60 : 0}
                outerRadius={80}
                paddingAngle={type === 'donut' ? 2 : 0}
                labelLine={false}
              >
                {data.map((index: number) => (
                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
            </PieChart>
          </ResponsiveContainer>
        );
      default:
        return (
          <div className="p-3 text-xs text-zinc-500 bg-zinc-50 rounded-lg">
            Unsupported chart type: {type}
          </div>
        );
    }
  };

  return (
    <details className="flex flex-col gap-2 my-4 border border-zinc-200/80 rounded-xl overflow-hidden group shadow-xs">
      <summary className="px-3.5 py-2 text-[12px] font-semibold text-zinc-600 cursor-pointer select-none bg-zinc-50/80 hover:bg-zinc-100 flex items-center gap-2 transition-colors">
        <svg className="w-3.5 h-3.5 text-zinc-400 group-open:rotate-90 transition-transform duration-200" viewBox="0 0 16 16" fill="none" stroke="currentColor">
          <path d="M6 12L10 8L6 4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
        <span className="flex-1">View Chart: {title || 'Preview'}</span>
      </summary>
      <div className="border-t border-zinc-100 p-4 bg-white">
        {renderChart()}
      </div>
    </details>
  );
}
