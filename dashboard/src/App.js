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

function App() {
  const [employees, setEmployees] = useState([]);
  
  // Preserved structural KPI states
  const [monthlyTargets, setMonthlyTargets] = useState({
    ticketsTakenPerAgent: "N/A",
    ticketsSolvedPerAgent: "N/A",
    totalTeamTakenTickets: "N/A",
    totalTeamSolvedTickets: "N/A",
    l2EscalationRate: "N/A"
  });

  const [summary, setSummary] = useState({
    weeklyHandleTarget: [0, 0, 0, 0],
    weeklyHandleActual: [0, 0, 0, 0],
    weeklySolveTarget: [0, 0, 0, 0],
    weeklySolveActual: [0, 0, 0, 0],
  });

  const API_BASE_URL = window.location.origin;

  // ==========================================
  // COMPREHENSIVE RECONCILIATION PARSER
  // ==========================================
  const parseExcelData = useCallback((arrayBuffer) => {
    setEmployees([]);

    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    
    // --- 1. PARSE SHEET: DASHBOARD (MONTHLY TARGETS & AGENTS) ---
    const dashboardSheetName = workbook.SheetNames.find((name) =>
      name.toLowerCase().includes("dashboard")
    ) || workbook.SheetNames[0];

    const dashboardSheet = workbook.Sheets[dashboardSheetName];
    const dashboardRows = XLSX.utils.sheet_to_json(dashboardSheet, { header: 1, defval: "" });

    let headerRowIndex = -1;
    const targetHolder = {
      ticketsTakenPerAgent: "N/A",
      ticketsSolvedPerAgent: "N/A",
      totalTeamTakenTickets: "N/A",
      totalTeamSolvedTickets: "N/A",
      l2EscalationRate: "N/A"
    };

    for (let i = 0; i < dashboardRows.length; i++) {
      const row = dashboardRows[i];
      if (!row || row.length === 0) continue;

      const cellA = row[0] ? String(row[0]).toLowerCase().trim() : "";
      const valB = row[1] !== undefined ? String(row[1]).trim() : "";

      if (cellA.includes("tickets taken per agent")) targetHolder.ticketsTakenPerAgent = valB;
      if (cellA.includes("tickets solved per agent")) targetHolder.ticketsSolvedPerAgent = valB;
      if (cellA.includes("total team taken tickets")) targetHolder.totalTeamTakenTickets = valB;
      if (cellA.includes("total team solved tickets")) targetHolder.totalTeamSolvedTickets = valB;
      if (cellA.includes("l2 escalation rate")) targetHolder.l2EscalationRate = valB;

      const containsAgentHeader = row.some(
        (cell) => String(cell).toLowerCase().trim() === "agent"
      );
      if (containsAgentHeader && headerRowIndex === -1) {
        headerRowIndex = i;
      }
    }
    setMonthlyTargets(targetHolder);

    let formattedAgents = [];
    if (headerRowIndex !== -1) {
      const headers = dashboardRows[headerRowIndex].map((h) => String(h).toLowerCase().trim());
      const agentIdx = headers.indexOf("agent");
      const ownedIdx = headers.findIndex((h) => h.includes("owned") || h.includes("new"));
      const escalatedIdx = headers.findIndex((h) => h.includes("escalated") || h.includes("l2"));
      const closedIdx = headers.findIndex((h) => h.includes("solved") || h.includes("closed"));
      const unresolvedIdx = headers.findIndex((h) => h.includes("unresolved"));

      for (let i = headerRowIndex + 1; i < dashboardRows.length; i++) {
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

    // --- 2. PARSE SHEET: WEEKLY TARGETS SPLIT (CHARTS DATA) ---
    let targetHandleArr = [0, 0, 0, 0];
    let actualHandleArr = [0, 0, 0, 0];
    let targetSolveArr = [0, 0, 0, 0];
    let actualSolveArr = [0, 0, 0, 0];

    const weeklySheetName = workbook.SheetNames.find((name) =>
      name.toLowerCase().includes("weekly")
    );

    if (weeklySheetName) {
      const weeklySheet = workbook.Sheets[weeklySheetName];
      const weeklyRows = XLSX.utils.sheet_to_json(weeklySheet, { header: 1, defval: "" });

      // Track how many rows of each metric type we have processed sequentially
      let handleRowCount = 0;
      let solveRowCount = 0;

      for (let i = 0; i < weeklyRows.length; i++) {
        const row = weeklyRows[i];
        if (!row || row.length === 0) continue;

        const metricName = row[0] ? String(row[0]).toLowerCase().trim() : "";

        // Extract values from column arrays safely
        const rowValues = [
          parseFloat(row[1]) || 0,
          parseFloat(row[2]) || 0,
          parseFloat(row[3]) || 0,
          parseFloat(row[4]) || 0
        ];

        // Process "Tickets to Handle" matches sequentially
        if (metricName.includes("tickets to handle")) {
          handleRowCount++;
          if (handleRowCount === 1) {
            // First match = Planned Target Goals
            targetHandleArr = rowValues;
          } else if (handleRowCount === 2) {
            // Second match = Accomplished/Actual Metrics
            actualHandleArr = rowValues;
          }
        }

        // Process "Tickets to Solve" matches sequentially
        if (metricName.includes("tickets to solve") || metricName.includes("solve/close")) {
          solveRowCount++;
          if (solveRowCount === 1) {
            // First match = Planned Target Goals
            targetSolveArr = rowValues;
          } else if (solveRowCount === 2) {
            // Second match = Accomplished/Actual Metrics
            actualSolveArr = rowValues;
          }
        }
      }
    }

    // Force an update to state with verified indices
    setSummary({
      weeklyHandleTarget: targetHandleArr,
      weeklyHandleActual: actualHandleArr,
      weeklySolveTarget: targetSolveArr,
      weeklySolveActual: actualSolveArr,
    });
  }, []);

  // ==========================================
  // BACKGROUND SERVER SYNCHRONIZATION EFFICIENCY
  // ==========================================
  useEffect(() => {
    fetch(`${API_BASE_URL}/api/file`)
      .then((res) => {
        if (!res.ok) throw new Error("No initial deployment data found on server filesystem.");
        return res.arrayBuffer();
      })
      .then((arrayBuffer) => parseExcelData(arrayBuffer))
      .catch((err) => console.log("Mount file sync status check:", err.message));
  }, [parseExcelData, API_BASE_URL]);

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const localBuffer = await file.arrayBuffer();
    parseExcelData(localBuffer);

    const formData = new FormData();
    formData.append("file", file);

    fetch(`${API_BASE_URL}/api/upload`, {
      method: "POST",
      body: formData,
    })
      .then((res) => res.json())
      .then((data) => console.log("VM Permanent Write confirmation state:", data.message))
      .catch((err) => console.error("Could not broadcast workspace tracking modifications:", err));
  };

  // ==========================================
  // BAR CONFIG DATA MANAGEMENT PIPELINES
  // ==========================================
  const chartLabels = ["W1 (Handled)", "W1 (Solved)", "W2 (Handled)", "W2 (Solved)", "W3 (Handled)", "W3 (Solved)", "W4 (Handled)", "W4 (Solved)"];
  
  const combinedBarData = {
    labels: chartLabels,
    datasets: [
      {
        label: "Target Goal",
        data: [
          summary.weeklyHandleTarget[0], summary.weeklySolveTarget[0],
          summary.weeklyHandleTarget[1], summary.weeklySolveTarget[1],
          summary.weeklyHandleTarget[2], summary.weeklySolveTarget[2],
          summary.weeklyHandleTarget[3], summary.weeklySolveTarget[3],
        ],
        backgroundColor: "rgba(255, 255, 255, 0.12)",
        borderColor: "rgba(255, 255, 255, 0.3)",
        borderWidth: 1,
      },
      {
        label: "Accomplished / Actual",
        data: [
          summary.weeklyHandleActual[0], summary.weeklySolveActual[0],
          summary.weeklyHandleActual[1], summary.weeklySolveActual[1],
          summary.weeklyHandleActual[2], summary.weeklySolveActual[2],
          summary.weeklyHandleActual[3], summary.weeklySolveActual[3],
        ],
        backgroundColor: "rgba(59, 130, 246, 0.75)",
        borderColor: "#3b82f6",
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
        borderRadius: 4,
      },
      {
        label: "Escalated to L2",
        data: employees.map((emp) => emp.escalated),
        backgroundColor: "rgba(168, 85, 247, 0.85)",
        borderRadius: 4,
      },
      {
        label: "Solved / Closed",
        data: employees.map((emp) => emp.closed),
        backgroundColor: "rgba(16, 185, 129, 0.85)",
        borderRadius: 4,
      },
      {
        label: "Unresolved Tickets",
        data: employees.map((emp) => emp.unresolved),
        backgroundColor: "rgba(239, 68, 68, 0.85)",
        borderRadius: 4,
      },
    ],
  };

  return (
    <div className="dashboard-container">
      <div className="mesh-gradient"></div>
      <div className="dashboard">
        
        {/* HEADER */}
        <header className="glass-header">
          <div>
            <h1>CloudBees Operations Dashboard</h1>
            <p>Upload your team operations spreadsheet template</p>
          </div>
          <div>
            <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
          </div>
        </header>

        {/* MONTHLY TARGET METRICS CARDS */}
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

        {/* WEEKLY OPERATIONS CHART PANEL */}
        <section className="full-width-chart-row">
          <div className="glass-card chart-box">
            <h3>Weekly Operational Overview (Target vs Accomplished)</h3>
            <div className="chart-container">
              <Bar 
                data={combinedBarData} 
                options={{ 
                  responsive: true, 
                  maintainAspectRatio: false,
                  plugins: { legend: { labels: { color: '#f8fafc' } } },
                  scales: {
                    x: { ticks: { color: 'rgba(255,255,255,0.6)' }, grid: { display: false } },
                    y: { ticks: { color: 'rgba(255,255,255,0.6)' }, grid: { color: 'rgba(255,255,255,0.03)' } }
                  }
                }} 
              />
            </div>
          </div>
        </section>

        {/* INDIVIDUAL PERFORMANCE RECORDS HORIZONTAL CHART */}
        <section className="full-width-chart-row">
          <div className="glass-card chart-box">
            <h3>Individual Performance Breakdown</h3>
            <div className="agent-chart-container">
              {employees.length > 0 ? (
                <Bar 
                  data={agentChartData} 
                  options={{ 
                    responsive: true, 
                    maintainAspectRatio: false,
                    indexAxis: 'y',
                    plugins: { 
                      legend: { 
                        position: 'top',
                        labels: { color: '#f8fafc', padding: 18, font: { size: 12 } } 
                      } 
                    },
                    scales: {
                      x: { 
                        ticks: { color: 'rgba(255,255,255,0.6)' }, 
                        grid: { color: 'rgba(255,255,255,0.03)' },
                        title: { display: true, text: 'Volume Count Metrics', color: 'rgba(255,255,255,0.4)' }
                      },
                      y: { 
                        ticks: { color: '#f8fafc', font: { weight: '600', size: 12 } }, 
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