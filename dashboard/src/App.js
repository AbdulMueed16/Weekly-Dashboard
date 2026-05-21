import React, { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { Bar, Doughnut } from "react-chartjs-2";

import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  ArcElement,
  Tooltip,
  Legend,
  PointElement,
  LineElement,
} from "chart.js";

import "./App.css";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  ArcElement,
  Tooltip,
  Legend,
  PointElement,
  LineElement
);

function App() {
  const [employees, setEmployees] = useState([]);
  const [monthlyTargets, setMonthlyTargets] = useState([]);

  const [summary, setSummary] = useState({
    totalTickets: 0,
    totalClosed: 0,
    totalBacklog: 0,
    totalEscalated: 0,
    escalationRate: 0,
    weeklyHandleTarget: [0, 0, 0, 0],
    weeklyHandleActual: [0, 0, 0, 0],
    weeklySolveTarget: [0, 0, 0, 0],
    weeklySolveActual: [0, 0, 0, 0],
  });

  // Since you set up Nginx reverse proxy routing, relative paths route automatically
  const API_BASE_URL = window.location.origin;

  // ==========================================
  // EXTRACTED REUSABLE PARSER LOGIC
  // ==========================================
  const parseExcelData = useCallback((arrayBuffer) => {
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // --- PARSE SHEET: DASHBOARD ---
    const dashboardSheetName = workbook.SheetNames.find(name => 
      name.toLowerCase().includes("dashboard")
    ) || workbook.SheetNames[0];
    
    const dashboardSheet = workbook.Sheets[dashboardSheetName];
    const dashboardRows = XLSX.utils.sheet_to_json(dashboardSheet, { header: 1 });

    let headerRowIndex = -1;
    let extractedTargets = [];
    let inMonthlyTargets = false;

    for (let i = 0; i < dashboardRows.length; i++) {
      const row = dashboardRows[i];
      if (!row || row.length === 0) continue;

      const firstCell = row[0] ? String(row[0]).toLowerCase().trim() : "";

      if (firstCell.includes("monthly targets")) {
        inMonthlyTargets = true;
        continue;
      }

      if (inMonthlyTargets) {
        if (firstCell.includes("individual performance")) {
          inMonthlyTargets = false;
        } else if (row[0]) {
          const label = String(row[0]).trim();
          const value = row[1];
          if (label.toLowerCase() !== "metric") {
            extractedTargets.push({
              label: label.toUpperCase(),
              value: value
            });
          }
        }
      }

      const containsAgentHeader = row.some(
        (cell) => String(cell).toLowerCase().trim() === "agent"
      );
      if (containsAgentHeader && headerRowIndex === -1) {
        headerRowIndex = i;
      }
    }

    setMonthlyTargets(extractedTargets);

    // Parse Individual Performance Table
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
      setEmployees(formattedAgents);
    }

    // --- PARSE SHEET: WEEKLY TARGETS SPLIT ---
    let weeklyHandleTarget = [0, 0, 0, 0];
    let weeklyHandleActual = [0, 0, 0, 0];
    let weeklySolveTarget = [0, 0, 0, 0];
    let weeklySolveActual = [0, 0, 0, 0];

    const weeklySheetName = workbook.SheetNames.find(name => name.toLowerCase().includes("weekly"));
    if (weeklySheetName) {
      const weeklySheet = workbook.Sheets[weeklySheetName];
      const weeklyRows = XLSX.utils.sheet_to_json(weeklySheet, { header: 1 });
      let processingAccomplished = false;

      weeklyRows.forEach((row) => {
        if (!row || row.length === 0) return;
        const firstCell = String(row[0]).toLowerCase().trim();

        if (firstCell.includes("accomplished")) {
          processingAccomplished = true;
          return;
        }
        if (firstCell.includes("tickets to handle")) {
          const values = [Number(row[1]) || 0, Number(row[2]) || 0, Number(row[3]) || 0, Number(row[4]) || 0];
          if (!processingAccomplished) weeklyHandleTarget = values;
          else weeklyHandleActual = values;
        }
        if (firstCell.includes("tickets to solve")) {
          const values = [Number(row[1]) || 0, Number(row[2]) || 0, Number(row[3]) || 0, Number(row[4]) || 0];
          if (!processingAccomplished) weeklySolveTarget = values;
          else weeklySolveActual = values;
        }
      });
    }

    const totalTickets = formattedAgents.reduce((sum, emp) => sum + emp.owned, 0);
    const totalClosed = formattedAgents.reduce((sum, emp) => sum + emp.closed, 0);
    const totalEscalated = formattedAgents.reduce((sum, emp) => sum + emp.escalated, 0);
    const totalBacklog = formattedAgents.reduce((sum, emp) => sum + emp.unresolved, 0);

    setSummary({
      totalTickets,
      totalClosed,
      totalEscalated,
      totalBacklog,
      escalationRate: totalTickets > 0 ? ((totalEscalated / totalTickets) * 100).toFixed(1) : 0,
      weeklyHandleTarget,
      weeklyHandleActual,
      weeklySolveTarget,
      weeklySolveActual,
    });
  }, []);

  // ==========================================
  // EFFECT: Fetch Existing Workbook on Mount
  // ==========================================
  useEffect(() => {
    fetch(`${API_BASE_URL}/api/file`)
      .then((res) => {
        if (!res.ok) throw new Error("No operational workspace found stored on VM.");
        return res.arrayBuffer();
      })
      .then((arrayBuffer) => parseExcelData(arrayBuffer))
      .catch((err) => console.log("Initial server check completed:", err.message));
  }, [parseExcelData, API_BASE_URL]);

  // ==========================================
  // HANDLE MANUEL FILE UPLOAD & BACKUP
  // ==========================================
  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    // 1. Instantly parse locally for immediate screen rendering feedback
    const localBuffer = await file.arrayBuffer();
    parseExcelData(localBuffer);

    // 2. Multi-user pipeline: push copy up to Python Flask filesystem permanent uploads folder
    const formData = new FormData();
    formData.append("file", file);

    fetch(`${API_BASE_URL}/api/upload`, {
      method: "POST",
      body: formData,
    })
      .then((res) => res.json())
      .then((data) => console.log("VM file write response state:", data.message))
      .catch((err) => console.error("Could not broadcast or persist workspace to Nginx host:", err));
  };

  // ==========================================
  // CHART CONFIGURATIONS
  // ==========================================
  const labels = [
    "W1 (Handled)", "W1 (Solved)", 
    "W2 (Handled)", "W2 (Solved)", 
    "W3 (Handled)", "W3 (Solved)", 
    "W4 (Handled)", "W4 (Solved)"
  ];

  const combinedBarData = {
    labels,
    datasets: [
      {
        label: "Target Goal",
        data: [
          summary.weeklyHandleTarget[0], summary.weeklySolveTarget[0],
          summary.weeklyHandleTarget[1], summary.weeklySolveTarget[1],
          summary.weeklyHandleTarget[2], summary.weeklySolveTarget[2],
          summary.weeklyHandleTarget[3], summary.weeklySolveTarget[3],
        ],
        backgroundColor: "rgba(255, 255, 255, 0.15)",
        borderColor: "rgba(255, 255, 255, 0.4)",
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

  const doughnutData = {
    labels: ["Solved", "Escalated", "Unresolved"],
    datasets: [
      {
        data: [summary.totalClosed, summary.totalEscalated, summary.totalBacklog],
        backgroundColor: ["#10b981", "#3b82f6", "#ef4444"],
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
            <h1>CloudBees Dashboard</h1>
            <p>Upload the excel sheet</p>
          </div>
          <div>
            <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
          </div>
        </header>

        {/* 100% DYNAMIC KPI CARDS SECTIONS */}
        <section className="kpi-grid">
          {monthlyTargets.length > 0 ? (
            monthlyTargets.map((target, idx) => (
              <div key={idx} className="glass-card kpi-card">
                <p className="card-label">{target.label}</p>
                <h2>{target.value}</h2>
              </div>
            ))
          ) : (
            <div className="glass-card kpi-card dynamic-placeholder-card">
              <p className="card-label">Operational Targets Status</p>
              <h2>Await Upload</h2>
              <p className="placeholder-text">Please select and upload an Excel workbook to extract dynamic metrics.</p>
            </div>
          )}
        </section>

        {/* SIDE-BY-SIDE CHART VIEWPORT BLOCK */}
        <section className="charts-side-by-side-row">
          
          {/* LEFT SIDE: COMBINED BAR GRAPH */}
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
                    x: { ticks: { color: 'rgba(255,255,255,0.7)' }, grid: { display: false } },
                    y: { ticks: { color: 'rgba(255,255,255,0.7)' }, grid: { color: 'rgba(255,255,255,0.05)' } }
                  }
                }} 
              />
            </div>
          </div>

          {/* RIGHT SIDE: TEAM RESOLUTION DOUGHNUT GRAPH */}
          <div className="glass-card chart-box">
            <h3>Team Resolution Efficiency</h3>
            <div className="doughnut-container">
              <Doughnut 
                data={doughnutData} 
                options={{ 
                  responsive: true, 
                  maintainAspectRatio: false,
                  plugins: { legend: { position: 'bottom', labels: { color: '#f8fafc', padding: 15 } } }
                }} 
              />
            </div>
          </div>

        </section>

        {/* INDIVIDUAL PERFORMANCE RECORDS TABLE SECTION */}
        <section className="table-row-full-width">
          <div className="glass-card table-section">
            <h3>Individual Performance Records</h3>
            <div className="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Agent Name</th>
                    <th>New Tickets Owned</th>
                    <th>Escalated To L2</th>
                    <th>Solved / Closed</th>
                    <th>Unresolved Tickets</th>
                  </tr>
                </thead>
                <tbody>
                  {employees.map((emp, index) => (
                    <tr key={index}>
                      <td className="agent-name">{emp.name}</td>
                      <td>{emp.owned}</td>
                      <td className={emp.escalated > 2 ? "high-alert" : ""}>{emp.escalated || 0}</td>
                      <td className="success-cell">{emp.closed}</td>
                      <td className="warning-cell">{emp.unresolved}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </section>

      </div>
    </div>
  );
}

export default App;