import json
import os


def generate_dashboard_html(report):
    """
    Generates a self-contained HTML dashboard with embedded report data.
    The recipient opens the HTML file in a browser and sees an interactive
    dashboard with charts — no need to upload JSON separately.
    """

    report_for_json = json.loads(json.dumps(report, default=str))

    report_json = json.dumps(report_for_json, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="uk">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TechFlow — Dashboard ({report['period']})</title>
    <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
    <style>
        :root {{
            --bg-primary: #0a0e1a;
            --bg-card: #111827;
            --bg-card-hover: #1a2337;
            --border: #1e293b;
            --text-primary: #f1f5f9;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --accent-blue: #3b82f6;
            --accent-emerald: #10b981;
            --accent-amber: #f59e0b;
            --accent-rose: #f43f5e;
            --accent-violet: #8b5cf6;
            --accent-cyan: #06b6d4;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'DM Sans', sans-serif;
            background: var(--bg-primary);
            color: var(--text-primary);
            min-height: 100vh;
        }}
        body::before {{
            content: '';
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.03'/%3E%3C/svg%3E");
            pointer-events: none; z-index: 0;
        }}
        .glow {{ position: fixed; border-radius: 50%; filter: blur(120px); pointer-events: none; z-index: 0; }}
        .g1 {{ width: 400px; height: 400px; background: rgba(59,130,246,0.15); top: -100px; right: -100px; }}
        .g2 {{ width: 300px; height: 300px; background: rgba(16,185,129,0.15); bottom: -50px; left: -50px; }}
        .container {{ position: relative; z-index: 1; max-width: 1280px; margin: 0 auto; padding: 40px 24px; }}
        .header {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 48px; animation: fadeDown 0.6s ease-out; }}
        .header h1 {{
            font-family: 'Space Mono', monospace; font-size: 28px; font-weight: 700;
            background: linear-gradient(135deg, var(--text-primary), var(--accent-blue));
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }}
        .header .sub {{ font-size: 14px; color: var(--text-muted); margin-top: 4px; font-family: 'Space Mono', monospace; }}
        .badges {{ display: flex; gap: 12px; align-items: center; }}
        .badge {{
            padding: 6px 14px; border-radius: 20px; font-size: 12px;
            font-family: 'Space Mono', monospace;
        }}
        .badge-live {{
            background: rgba(16,185,129,0.1); border: 1px solid rgba(16,185,129,0.2); color: var(--accent-emerald);
            display: flex; align-items: center; gap: 6px;
        }}
        .badge-live .dot {{ width: 6px; height: 6px; border-radius: 50%; background: var(--accent-emerald); animation: pulse 2s infinite; }}
        .badge-period {{ background: rgba(59,130,246,0.1); border: 1px solid rgba(59,130,246,0.2); color: var(--accent-blue); }}
        .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 32px; }}
        .card {{
            background: var(--bg-card); border: 1px solid var(--border); border-radius: 16px;
            padding: 24px; transition: all 0.3s; animation: fadeUp 0.6s ease-out backwards;
        }}
        .card:hover {{ background: var(--bg-card-hover); transform: translateY(-2px); border-color: var(--accent-blue); box-shadow: 0 8px 32px rgba(59,130,246,0.1); }}
        .card:nth-child(1){{animation-delay:.1s}} .card:nth-child(2){{animation-delay:.15s}}
        .card:nth-child(3){{animation-delay:.2s}} .card:nth-child(4){{animation-delay:.25s}}
        .card:nth-child(5){{animation-delay:.3s}}
        .card-label {{ font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: var(--text-muted); margin-bottom: 8px; }}
        .card-value {{ font-family: 'Space Mono', monospace; font-size: 36px; font-weight: 700; line-height: 1; }}
        .card-unit {{ font-size: 14px; color: var(--text-secondary); margin-top: 4px; }}
        .v-blue {{ color: var(--accent-blue); }} .v-emerald {{ color: var(--accent-emerald); }}
        .v-amber {{ color: var(--accent-amber); }} .v-violet {{ color: var(--accent-violet); }}
        .v-rose {{ color: var(--accent-rose); }}
        .charts {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 32px; }}
        .chart-card {{ background: var(--bg-card); border: 1px solid var(--border); border-radius: 16px; padding: 24px; animation: fadeUp 0.8s ease-out backwards; }}
        .chart-card:nth-child(1){{animation-delay:.35s}} .chart-card:nth-child(2){{animation-delay:.4s}}
        .chart-title {{ font-size: 14px; text-transform: uppercase; letter-spacing: 1px; color: var(--text-muted); margin-bottom: 20px; }}
        .chart-box {{ position: relative; height: 260px; }}
        .tbl-card {{ background: var(--bg-card); border: 1px solid var(--border); border-radius: 16px; padding: 24px; animation: fadeUp 0.8s ease-out backwards; animation-delay: .45s; margin-bottom: 32px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ text-align: left; padding: 12px 16px; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: var(--text-muted); border-bottom: 1px solid var(--border); }}
        td {{ padding: 14px 16px; font-size: 14px; border-bottom: 1px solid rgba(30,41,59,0.5); }}
        tr:last-child td {{ border-bottom: none; }}
        tr:hover td {{ background: rgba(59,130,246,0.03); }}
        .name-cell {{ display: flex; align-items: center; gap: 10px; font-weight: 500; }}
        .avatar {{
            width: 32px; height: 32px; border-radius: 8px; display: flex;
            align-items: center; justify-content: center; font-size: 14px; font-weight: 700; color: #fff;
        }}
        .bar-cell {{ display: flex; align-items: center; gap: 10px; }}
        .bar-track {{ flex: 1; height: 6px; background: rgba(255,255,255,0.05); border-radius: 3px; overflow: hidden; }}
        .bar-fill {{ height: 100%; border-radius: 3px; transition: width 1.5s cubic-bezier(0.4,0,0.2,1); }}
        .bar-val {{ font-family: 'Space Mono', monospace; font-size: 14px; font-weight: 700; min-width: 24px; text-align: right; }}
        .footer {{ text-align: center; padding: 32px 0; color: var(--text-muted); font-size: 12px; font-family: 'Space Mono', monospace; }}
        @keyframes fadeDown {{ from {{ opacity:0; transform:translateY(-20px) }} to {{ opacity:1; transform:translateY(0) }} }}
        @keyframes fadeUp {{ from {{ opacity:0; transform:translateY(20px) }} to {{ opacity:1; transform:translateY(0) }} }}
        @keyframes pulse {{ 0%,100% {{ opacity:1 }} 50% {{ opacity:0.4 }} }}
        @media (max-width:768px) {{ .charts {{ grid-template-columns:1fr; }} .grid {{ grid-template-columns:repeat(2,1fr); }} .header {{ flex-direction:column; gap:16px; }} }}
    </style>
</head>
<body>
<div class="glow g1"></div>
<div class="glow g2"></div>
<div class="container">
    <div class="header">
        <div>
            <h1>TechFlow Dashboard</h1>
            <div class="sub" id="period"></div>
        </div>
        <div class="badges">
            <div class="badge badge-live"><span class="dot"></span> Report</div>
            <div class="badge badge-period" id="date"></div>
        </div>
    </div>
    <div class="grid">
        <div class="card"><div class="card-label">Нових запитів</div><div class="card-value v-blue" id="m-new">0</div><div class="card-unit">за тиждень</div></div>
        <div class="card"><div class="card-label">Закритих запитів</div><div class="card-value v-emerald" id="m-closed">0</div><div class="card-unit">за тиждень</div></div>
        <div class="card"><div class="card-label">Час обробки</div><div class="card-value v-amber" id="m-proc">0</div><div class="card-unit">годин (середній)</div></div>
        <div class="card"><div class="card-label">Час реакції</div><div class="card-value v-violet" id="m-react">0</div><div class="card-unit">годин (середній)</div></div>
        <div class="card"><div class="card-label">Прострочені</div><div class="card-value v-rose" id="m-over">0</div><div class="card-unit">запитів &gt;24г</div></div>
    </div>
    <div class="charts">
        <div class="chart-card"><div class="chart-title">📂 Розподіл по послугах</div><div class="chart-box"><canvas id="c-svc"></canvas></div></div>
        <div class="chart-card"><div class="chart-title">📋 Навантаження консультантів</div><div class="chart-box"><canvas id="c-wl"></canvas></div></div>
    </div>
    <div class="tbl-card">
        <div class="chart-title">🏆 Консультанти — відкриті запити</div>
        <table><thead><tr><th>Консультант</th><th>Відкритих</th><th style="width:50%">Навантаження</th></tr></thead><tbody id="tbl"></tbody></table>
    </div>
    <div class="footer">TechFlow Consulting — автоматично згенерований дашборд</div>
</div>
<script>
const DATA = {report_json};
const COLORS = ['#3b82f6','#10b981','#f59e0b','#f43f5e','#8b5cf6','#06b6d4'];
const AVATARS = ['linear-gradient(135deg,#3b82f6,#8b5cf6)','linear-gradient(135deg,#10b981,#06b6d4)','linear-gradient(135deg,#f59e0b,#f43f5e)','linear-gradient(135deg,#8b5cf6,#ec4899)'];
Chart.defaults.color='#94a3b8';
Chart.defaults.font.family="'DM Sans',sans-serif";
Chart.defaults.plugins.legend.labels.padding=16;
Chart.defaults.plugins.legend.labels.usePointStyle=true;
Chart.defaults.plugins.legend.labels.pointStyle='circle';

function anim(el,t,d=1000){{const s=performance.now(),f=String(t).includes('.');function u(c){{const p=Math.min((c-s)/d,1),e=1-Math.pow(1-p,3),v=t*e;el.textContent=f?v.toFixed(2):Math.round(v);if(p<1)requestAnimationFrame(u)}}requestAnimationFrame(u)}}

const m=DATA.metrics;
document.getElementById('period').textContent=DATA.period;
document.getElementById('date').textContent=DATA.report_date;
anim(document.getElementById('m-new'),m.new_requests_this_week);
anim(document.getElementById('m-closed'),m.closed_requests_this_week);
anim(document.getElementById('m-proc'),m.avg_processing_time_hours,1200);
anim(document.getElementById('m-react'),m.avg_reaction_time_hours,1200);
anim(document.getElementById('m-over'),m.overdue_count);

const svcL=m.service_stats.map(s=>s.service), svcD=m.service_stats.map(s=>s.count);
new Chart(document.getElementById('c-svc'),{{type:'doughnut',data:{{labels:svcL,datasets:[{{data:svcD,backgroundColor:COLORS.slice(0,svcL.length),borderWidth:0,hoverOffset:8}}]}},options:{{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{{legend:{{position:'right',labels:{{font:{{size:12}},padding:12}}}}}},animation:{{animateRotate:true,duration:1200}}}}}});

const cons=Object.keys(m.consultant_workload), wlD=Object.values(m.consultant_workload);
new Chart(document.getElementById('c-wl'),{{type:'bar',data:{{labels:cons,datasets:[{{label:'Відкритих запитів',data:wlD,backgroundColor:COLORS.slice(0,cons.length).map(c=>c+'33'),borderColor:COLORS.slice(0,cons.length),borderWidth:2,borderRadius:8,borderSkipped:false}}]}},options:{{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{{legend:{{display:false}}}},scales:{{x:{{grid:{{color:'rgba(255,255,255,0.04)'}},ticks:{{stepSize:1}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:13,weight:500}}}}}}}},animation:{{duration:1200}}}}}});

const mx=Math.max(...wlD,1),tb=document.getElementById('tbl');
cons.forEach((n,i)=>{{const c=wlD[i],p=c/mx*100,cl=COLORS[i%COLORS.length],av=AVATARS[i%AVATARS.length];
tb.innerHTML+=`<tr><td><div class="name-cell"><div class="avatar" style="background:${{av}}">${{n[0].toUpperCase()}}</div>${{n}}</div></td><td style="font-family:'Space Mono',monospace;font-weight:700;color:${{cl}}">${{c}}</td><td><div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:0%;background:${{cl}}" data-w="${{p}}%"></div></div><div class="bar-val" style="color:${{cl}}">${{c}}</div></div></td></tr>`}});
setTimeout(()=>document.querySelectorAll('.bar-fill').forEach(b=>b.style.width=b.dataset.w),100);
</script>
</body>
</html>"""

    return html


def save_dashboard(dashboard_html, directory, filename):
    """Save dashboard HTML file."""
    filepath = os.path.join(directory, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(dashboard_html)
    print(f"📊 Dashboard збережено: {filepath}")
    return filepath
