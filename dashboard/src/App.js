import React, { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { Bar } from "react-chartjs-2";

import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Tooltip,
  Legend,
} from "chart.js";

import "./App.css";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Tooltip,
  Legend
);

// =========================================================================
// EXECUTIVE PLUG-IN: STAMPS METRIC VALUES PRECISELY INSIDE THE BAR SEGMENTS
// =========================================================================
const chartValueLabelsPlugin = {
  id: "chartValueLabels",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    ctx.save();
    ctx.font = "bold 11px 'Inter', sans-serif";
    ctx.fillStyle = "#f8fafc"; 
    
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      meta.data.forEach((element, index) => {
        const value = dataset.data[index];
        if (value === undefined || value === null || value === 0) return;

        if (chart.options.scales?.x?.stacked && chart.options.indexAxis === "y") {
          const midpointX = (element.x + element.base) / 2;
          const barWidth = Math.abs(element.x - element.base);

          if (barWidth > 16) {
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillText(value, midpointX, element.y);
          }
        } else {
          ctx.textAlign = "center";
          ctx.textBaseline = "bottom";
          ctx.fillText(value, element.x, element.y - 4);
        }
      });
    });
    ctx.restore();
  }
};

// =========================================================================
// EXECUTIVE PLUG-IN: ADDS RESPONSIVE SPACE BETWEEN LEGEND AND GRAPH AREA
// =========================================================================
const legendMarginPlugin = {
  id: "legendMargin",
  beforeInit(chart) {
    const originalFit = chart.legend.fit;
    chart.legend.fit = function fit() {
      if (originalFit) {
        originalFit.bind(chart.legend)();
      }
      this.height += 20; 
    };
  }
};

function App() {
  const [employees, setEmployees] = useState([]);
  
  // Preserved Monthly KPI Targets States
  const [monthlyTargets, setMonthlyTargets] = useState({
    ticketsTakenPerAgent: "N/A",
    ticketsSolvedPerAgent: "N/A",
    totalTeamTakenTickets: "N/A",
    totalTeamSolvedTickets: "N/A",
    l2EscalationRate: "N/A",
  });

  // Dynamic States for Separated Latest Weekly Highlight Metrics
  const [latestWeekLabel, setLatestWeekLabel] = useState("Latest Week");
  const [latestWeeklyMetrics, setLatestWeeklyMetrics] = useState({
    l2Escalated: "N/A",
    unresolvedBacklog: "N/A"
  });

  // Dynamic Chart Header Labels from Spreadsheet
  const [targetLabels, setTargetLabels] = useState(["Week 1", "Week 2", "Week 3", "Week 4"]);
  const [actualLabels, setActualLabels] = useState(["Week 1", "Week 2", "Week 3", "Week 4"]);

  // Highly-Isolated Weekly Charts Data States
  const [weeklyTargetData, setWeeklyTargetData] = useState({
    handle: [0, 0, 0, 0],
    solve: [0, 0, 0, 0]
  });

  const [weeklyActualData, setWeeklyActualData] = useState({
    handle: [0, 0, 0, 0],
    solve: [0, 0, 0, 0]
  });

  const API_BASE_URL = window.location.origin;

  // =========================================================
  // SINGLE-SHEET HIGH FIDELITY MATRIX EXCEL PARSER
  // =========================================================
  const parseExcelData = useCallback((arrayBuffer) => {
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // Target ONLY the master dashboard worksheet context explicitly
    const dashboardSheetName = workbook.SheetNames.find((name) =>
      name.toLowerCase().includes("dashboard")
    ) || workbook.SheetNames[0];

    const dashboardSheet = workbook.Sheets[dashboardSheetName];
    const dashboardRows = XLSX.utils.sheet_to_json(dashboardSheet, { header: 1, defval: "" });

    let agentHeaderIdx = -1;
    let isAccomplishedBlock = false; 

    let tempTargetLabels = [];
    let tempActualLabels = [];
    let tHandle = [0, 0, 0, 0];
    let tSolve = [0, 0, 0, 0];
    let aHandle = [0, 0, 0, 0];
    let aSolve = [0, 0, 0, 0];

    // Holders to isolate latest weeks metrics via clean back-scanning
    let accomplishedWeekHeaders = ["Week 1", "Week 2", "Week 3", "Week 4"];
    let rawL2EscalatedRow = ["", "", "", ""];
    let rawUnresolvedRow = ["", "", "", ""];

    const targetHolder = {
      ticketsTakenPerAgent: "N/A",
      ticketsSolvedPerAgent: "N/A",
      totalTeamTakenTickets: "N/A",
      totalTeamSolvedTickets: "N/A",
      l2EscalationRate: "N/A",
    };

    for (let i = 0; i < dashboardRows.length; i++) {
      const row = dashboardRows[i];
      if (!row || row.length === 0) continue;

      const cellA = row[0] ? String(row[0]).toLowerCase().trim() : "";
      const valB = row[1] !== undefined ? String(row[1]).trim() : "";

      // 1. Extract Top-Level Monthly Targets Overwrites
      if (cellA.includes("tickets taken per agent")) targetHolder.ticketsTakenPerAgent = valB;
      if (cellA.includes("tickets solved per agent")) targetHolder.ticketsSolvedPerAgent = valB;
      if (cellA.includes("total team taken tickets")) targetHolder.totalTeamTakenTickets = valB;
      if (cellA.includes("total team solved tickets")) targetHolder.totalTeamSolvedTickets = valB;
      if (cellA.includes("l2 escalation rate")) targetHolder.l2EscalationRate = valB;

      // 2. Detect Boundary State Switch to the Accomplished Blocks
      if (cellA.includes("accomplished")) {
        isAccomplishedBlock = true;
        continue;
      }

      // 3. Dynamically capture column header strings as they appear in the file
      if (cellA === "metric") {
        const labels = [
          row[1] ? String(row[1]).trim() : "W1",
          row[2] ? String(row[2]).trim() : "W2",
          row[3] ? String(row[3]).trim() : "W3",
          row[4] ? String(row[4]).trim() : "W4",
        ];
        if (!isAccomplishedBlock) {
          tempTargetLabels = labels;
        } else {
          tempActualLabels = labels;
          accomplishedWeekHeaders = labels;
        }
        continue;
      }

      // 4. Extract data rows for targets and actuals from Dashboard sheet
      const values = [
        parseFloat(row[1]) || 0,
        parseFloat(row[2]) || 0,
        parseFloat(row[3]) || 0,
        parseFloat(row[4]) || 0,
      ];

      if (cellA.includes("tickets to handle") || cellA.includes("tickets to handled")) {
        if (!isAccomplishedBlock) tHandle = values;
        else aHandle = values;
      }

      if (cellA.includes("tickets to solve") || cellA.includes("solved/closed") || cellA.includes("solved")) {
        if (!isAccomplishedBlock) tSolve = values;
        else aSolve = values;
      }

      // 5. Populate intermediate arrays for back-scanning latest weeks records
      if (isAccomplishedBlock && cellA.includes("l2 escalated")) {
        rawL2EscalatedRow = [row[1], row[2], row[3], row[4]].map(v => v !== undefined ? String(v).trim() : "");
      }
      if (isAccomplishedBlock && (cellA.includes("unresolved tickets") || cellA.includes("backlog"))) {
        rawUnresolvedRow = [row[1], row[2], row[3], row[4]].map(v => v !== undefined ? String(v).trim() : "");
      }

      // 6. Locate Lower Agent Grid Header Index Boundary Breakpoint
      const containsAgentHeader = row.some(cell => String(cell).toLowerCase().trim() === "agent");
      if (containsAgentHeader && agentHeaderIdx === -1) {
        agentHeaderIdx = i;
      }
    }

    // --- REVERSE SCANNING RESOLUTION ENGINE FOR LATEST ACTIVE WEEK ---
    let latestActiveWeekIdx = 0;
    for (let w = 3; w >= 0; w--) {
      if (rawL2EscalatedRow[w] !== "" || rawUnresolvedRow[w] !== "") {
        latestActiveWeekIdx = w;
        break;
      }
    }

    // Set high-fidelity isolated metrics states
    setMonthlyTargets(targetHolder);
    if (tempTargetLabels.length > 0) setTargetLabels(tempTargetLabels);
    if (tempActualLabels.length > 0) setActualLabels(tempActualLabels);

    setWeeklyTargetData({ handle: tHandle, solve: tSolve });
    setWeeklyActualData({ handle: aHandle, solve: aSolve });

    // Assign isolated latest week metrics values safely
    setLatestWeekLabel(accomplishedWeekHeaders[latestActiveWeekIdx] || "Latest Week");
    setLatestWeeklyMetrics({
      l2Escalated: rawL2EscalatedRow[latestActiveWeekIdx] || "0",
      unresolvedBacklog: rawUnresolvedRow[latestActiveWeekIdx] || "0"
    });

    // --- 7. PARSE AGENT PERFORMANCE GRID ---
    let formattedAgents = [];
    if (agentHeaderIdx !== -1) {
      const headers = dashboardRows[agentHeaderIdx].map(h => String(h).toLowerCase().trim());
      const agentIdx = headers.indexOf("agent");
      const ownedIdx = headers.findIndex(h => h.includes("owned") || h.includes("new"));
      const escalatedIdx = headers.findIndex(h => h.includes("escalated") || h.includes("l2"));
      const closedIdx = headers.findIndex(h => h.includes("solved") || h.includes("closed"));
      const unresolvedIdx = headers.findIndex(h => h.includes("unresolved"));

      for (let i = agentHeaderIdx + 1; i < dashboardRows.length; i++) {
        const row = dashboardRows[i];
        if (!row || row.length === 0) continue;

        const agentName = row[agentIdx];
        if (!agentName || String(agentName).trim() === "" || String(agentName).toLowerCase().includes("total")) {
          continue;
        }

        formattedAgents.push({
          name: String(agentName).trim(),
          owned: Number(row[ownedIdx] || 0),
          escalated: Number(row[escalatedIdx] || 0),
          closed: Number(row[closedIdx] || 0),
          unresolved: Number(row[unresolvedIdx] || 0),
        });
      }
    }
    setEmployees(formattedAgents);
  }, []);

  useEffect(() => {
    fetch(`${API_BASE_URL}/api/file`)
      .then((res) => {
        if (!res.ok) throw new Error("No operational dataset mapped on server.");
        return res.arrayBuffer();
      })
      .then((arrayBuffer) => parseExcelData(arrayBuffer))
      .catch((err) => console.log("Mount file check sync:", err.message));
  }, [parseExcelData, API_BASE_URL]);

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    setWeeklyTargetData({ handle: [0, 0, 0, 0], solve: [0, 0, 0, 0] });
    setWeeklyActualData({ handle: [0, 0, 0, 0], solve: [0, 0, 0, 0] });
    setEmployees([]);

    const localBuffer = await file.arrayBuffer();
    parseExcelData(localBuffer);

    const formData = new FormData();
    formData.append("file", file);

    fetch(`${API_BASE_URL}/api/upload`, { method: "POST", body: formData })
      .then((res) => res.json())
      .then((data) => console.log("Server workbook write success:", data.message))
      .catch((err) => console.error("Could not sync payload upload:", err));
  };

  const chartLabels = ["W1 (Handled)", "W1 (Solved)", "W2 (Handled)", "W2 (Solved)", "W3 (Handled)", "W3 (Solved)", "W4 (Handled)", "W4 (Solved)"];
  
  const targetChartData = {
    labels: targetLabels,
    datasets: [
      {
        label: "Tickets to Handle",
        data: weeklyTargetData.handle,
        backgroundColor: "rgba(255, 255, 255, 0.12)",
        borderColor: "rgba(255, 255, 255, 0.3)",
        borderWidth: 1,
      },
      {
        label: "Tickets to Solve / Close",
        data: weeklyTargetData.solve,
        backgroundColor: "rgba(59, 130, 246, 0.75)",
        borderColor: "#3b82f6",
        borderWidth: 1,
      },
    ],
  };

  const actualChartData = {
    labels: actualLabels,
    datasets: [
      {
        label: "Tickets Handled",
        data: weeklyActualData.handle,
        backgroundColor: "rgba(16, 185, 129, 0.2)",
        borderColor: "rgba(16, 185, 129, 0.6)",
        borderWidth: 1,
      },
      {
        label: "Tickets Solved / Closed",
        data: weeklyActualData.solve,
        backgroundColor: "rgba(16, 185, 129, 0.85)",
        borderColor: "#10b981",
        borderWidth: 1,
      },
    ],
  };

  const agentLabels = employees.map((emp) => emp.name);
  const agentChartData = {
    labels: agentLabels,
    datasets: [
      {
        label: "New Tickets Owned",
        data: employees.map((emp) => emp.owned),
        backgroundColor: "rgba(59, 130, 246, 0.85)",
      },
      {
        label: "Escalated to L2",
        data: employees.map((emp) => emp.escalated),
        backgroundColor: "rgba(168, 85, 247, 0.85)",
      },
      {
        label: "Solved / Closed",
        data: employees.map((emp) => emp.closed),
        backgroundColor: "rgba(16, 185, 129, 0.85)",
      },
      {
        label: "Unresolved Tickets",
        data: employees.map((emp) => emp.unresolved),
        backgroundColor: "rgba(239, 68, 68, 0.85)",
      },
    ],
  };

  const standardChartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { 
      legend: { labels: { color: '#f8fafc', font: { size: 11 } } },
      tooltip: { enabled: false }
    },
    hover: { mode: null }, 
    scales: {
      x: { ticks: { color: 'rgba(255,255,255,0.6)', font: { size: 10 } }, grid: { display: false } },
      y: { ticks: { color: 'rgba(255,255,255,0.6)' }, grid: { color: 'rgba(255,255,255,0.03)' } }
    }
  };

  return (
    <div className="dashboard-container">
      <div className="mesh-gradient"></div>
      <div className="dashboard">
        
        {/* HEADER PANEL */}
        <header className="glass-header">
          <div>
            <h1>CloudBees Team Targets & Performance Dashboard</h1>
            <p>Upload your team operations spreadsheet template</p>
          </div>
          <div>
            <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
          </div>
        </header>

        {/* COMBINED TARGETS AND SEPARATED LATEST WEEK METRICS SPLIT ASSEMBLY */}
        <div className="metrics-master-split">
          <div className="metrics-group-left">
            <h3 className="corporate-row-label">Monthly Targets Overview</h3>
            <section className="kpi-row-layout">
              <div className="glass-card matrix-kpi-box">
                <p className="card-label">Tickets taken per agent</p>
                <h2>{monthlyTargets.ticketsTakenPerAgent}</h2>
              </div>
              <div className="glass-card matrix-kpi-box">
                <p className="card-label">Tickets solved per agent</p>
                <h2>{monthlyTargets.ticketsSolvedPerAgent}</h2>
              </div>
              <div className="glass-card matrix-kpi-box">
                <p className="card-label">Total team taken tickets</p>
                <h2>{monthlyTargets.totalTeamTakenTickets}</h2>
              </div>
              <div className="glass-card matrix-kpi-box">
                <p className="card-label">Total team solved tickets</p>
                <h2>{monthlyTargets.totalTeamSolvedTickets}</h2>
              </div>
              <div className="glass-card matrix-kpi-box">
                <p className="card-label">L2 escalation rate</p>
                <h2>{monthlyTargets.l2EscalationRate}</h2>
              </div>
            </section>
          </div>

          <div className="metrics-group-right">
            <h3 className="corporate-row-label blue-bar">Accomplished: {latestWeekLabel}</h3>
            <section className="kpi-row-layout split-two">
              <div className="glass-card matrix-kpi-box purple-accent">
                <p className="card-label">L2 Escalated</p>
                <h2>{latestWeeklyMetrics.l2Escalated}</h2>
              </div>
              <div className="glass-card matrix-kpi-box red-accent">
                <p className="card-label">Total Unresolved Tickets [Backlog]</p>
                <h2>{latestWeeklyMetrics.unresolvedBacklog}</h2>
              </div>
            </section>
          </div>
        </div>

        {/* SIDE-BY-SIDE DOUBLE CHART SEGMENTS ROW */}
        <section className="charts-side-by-side-row">
          <div className="glass-card chart-box">
            <h3>Weekly Operational Planned Targets</h3>
            <div className="chart-container">
              <Bar data={targetChartData} options={standardChartOptions} plugins={[chartValueLabelsPlugin, legendMarginPlugin]} />
            </div>
          </div>

          <div className="glass-card chart-box">
            <h3>Weekly Accomplished Performance Metrics</h3>
            <div className="chart-container">
              <Bar data={actualChartData} options={standardChartOptions} plugins={[chartValueLabelsPlugin, legendMarginPlugin]} />
            </div>
          </div>
        </section>

        {/* PROFESSIONAL STACKED HORIZONTAL PERFORMANCE GRAPH */}
        <section className="full-width-chart-row">
          <div className="glass-card chart-box">
            <h3>Individual Performance Breakdown</h3>
            <div className="agent-chart-container">
              {employees.length > 0 ? (
                <Bar 
                  data={agentChartData} 
                  plugins={[chartValueLabelsPlugin, legendMarginPlugin]} 
                  options={{ 
                    responsive: true, 
                    maintainAspectRatio: false,
                    indexAxis: 'y', 
                    plugins: { 
                      legend: { 
                        position: 'top',
                        labels: { color: '#f8fafc', font: { size: 12, weight: '600' } } 
                      },
                      tooltip: { enabled: false }
                    },
                    hover: { mode: null },
                    scales: {
                      x: { 
                        stacked: true, 
                        ticks: { color: 'rgba(255,255,255,0.6)' }, 
                        grid: { color: 'rgba(255,255,255,0.03)' },
                        title: { display: true, text: 'Total Aggregated Tickets Workload Contribution', color: 'rgba(255,255,255,0.4)' }
                      },
                      y: { 
                        stacked: true, 
                        ticks: { color: '#f8fafc', font: { weight: '600', size: 11 } }, 
                        grid: { display: false } 
                      }
                    }
                  }} 
                />
              ) : (
                <div className="center-empty">Please upload an operational spreadsheet template to view agent metrics.</div>
              )}
            </div>
          </div>
        </section>

      </div>
    </div>
  );
}

export default App;