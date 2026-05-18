import React, { useEffect, useState } from "react";
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
  const [summary, setSummary] = useState({
    totalTickets: 0,
    totalClosed: 0,
    totalBacklog: 0,
    totalEscalated: 0,
    escalationRate: 0,
    targetTickets: 171,      // Default fallback values
    targetClosed: 135,
    targetEscalation: "< 30%",
    weeklyTarget: 42,
  });

  const EXCEL_FILE_URL = "/data/dashboard.xlsx";

  const fetchExcelData = async () => {
    try {
      const response = await fetch(EXCEL_FILE_URL);
      if (!response.ok) throw new Error("Failed to load Excel file");

      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });

      // Target Sheet 1 (Index 0)
      const firstSheetName = workbook.SheetNames[0];
      const agentSheet = workbook.Sheets[firstSheetName];
      
      // Read sheet as a raw 2D grid matrix to parse multiple layouts safely
      const rawRows = XLSX.utils.sheet_to_json(agentSheet, { header: 1 });

      let headerRowIndex = -1;
      let extractedTargetTickets = 171;
      let extractedTargetClosed = 135;
      let extractedTargetEscalation = "< 30%";
      let extractedTargetWeekly = 42;

      // 1. Scan rows to extract Target Goals and locate the Agent Leaderboard headers
      for (let i = 0; i < rawRows.length; i++) {
        const row = rawRows[i];
        if (!row) continue;

        // Extract values from the Monthly Targets section (Columns 0 & 1)
        if (row[0]) {
          const cellText = String(row[0]).toLowerCase().trim();
          if (cellText.includes("total team taken tickets")) {
            extractedTargetTickets = row[1];
          } else if (cellText.includes("total team solved tickets")) {
            extractedTargetClosed = row[1];
          } else if (cellText.includes("l2 escalation rate")) {
            extractedTargetEscalation = row[1];
          }
        }

        // Extract values from the Weekly Targets section (Columns 3 & 4)
        if (row[3]) {
          const weeklyCellText = String(row[3]).toLowerCase().trim();
          if (weeklyCellText.includes("total tickets to handle")) {
            extractedTargetWeekly = row[4];
          }
        }

        // Pinpoint the row containing the "Agent" column header
        const containsAgentHeader = row.some(
          (cell) => String(cell).toLowerCase().trim() === "agent"
        );

        if (containsAgentHeader && headerRowIndex === -1) {
          headerRowIndex = i;
        }
      }

      // 2. Parse the Individual Performance Agent rows if header row was found
      if (headerRowIndex !== -1) {
        const headers = rawRows[headerRowIndex].map((h) =>
          String(h).toLowerCase().trim()
        );

        const agentIdx = headers.indexOf("agent");
        const ownedIdx = headers.findIndex((h) => h.includes("owned") || h.includes("new tickets"));
        const escalatedIdx = headers.findIndex((h) => h.includes("escalated") || h.includes("l2"));
        const closedIdx = headers.findIndex((h) => h.includes("solved") || h.includes("closed"));
        const unresolvedIdx = headers.findIndex((h) => h.includes("unresolved"));

        const formattedAgents = [];

        for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
          const row = rawRows[i];
          if (!row || row.length === 0) continue;

          const agentName = row[agentIdx];

          // Skip any footer summary fields or empty gap cells
          if (
            !agentName ||
            String(agentName).trim() === "" ||
            String(agentName).toLowerCase().includes("total")
          ) {
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

        // 3. Combine aggregated metrics with targets extracted from the sheet
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
          targetTickets: extractedTargetTickets,
          targetClosed: extractedTargetClosed,
          targetEscalation: extractedTargetEscalation,
          weeklyTarget: Number(extractedTargetWeekly) || 42,
        });

        console.log("Excel target structures and individual performance loaded successfully.");
      } else {
        console.error("Could not find individual performance table headers.");
      }

    } catch (error) {
      console.error("Local Excel Read Error:", error);
    }
  };

  useEffect(() => {
    fetchExcelData();
  }, []);

  // ==========================================
  // BAR CHART DYNAMIC CONFIGURATION
  // ==========================================
  const barChartData = {
    labels: ["Week 1", "Week 2", "Week 3", "Week 4"],
    datasets: [
      {
        label: "Actual Handled",
        data: [summary.totalTickets, 0, 0, 0],
        backgroundColor: "rgba(59, 130, 246, 0.6)",
        borderColor: "#3b82f6",
        borderWidth: 1,
        borderRadius: 5,
      },
      {
        label: "Target Goal",
        data: [summary.weeklyTarget, summary.weeklyTarget, summary.weeklyTarget, summary.weeklyTarget],
        backgroundColor: "rgba(255, 255, 255, 0.1)",
        borderColor: "rgba(255, 255, 255, 0.3)",
        borderWidth: 1,
        borderDash: [5, 5],
        borderRadius: 5,
      },
    ],
  };

  // ==========================================
  // DOUGHNUT CHART DYNAMIC CONFIGURATION
  // ==========================================
  const doughnutData = {
    labels: ["Solved", "Escalated", "Unresolved"],
    datasets: [
      {
        data: [summary.totalClosed, summary.totalEscalated, summary.totalBacklog],
        backgroundColor: ["#10b981", "#3b82f6", "rgba(239, 68, 68, 0.5)"],
        hoverOffset: 15,
        borderWidth: 0,
      },
    ],
  };

  return (
    <div className="dashboard-container">
      <div className="mesh-gradient"></div>

      <div className="dashboard">
        {/* HEADER */}
        <header className="glass-header animate-drop">
          <div className="brand">
            <div className="logo-sq"></div>
            <div>
              <h1>TICKET COMMAND</h1>
              <p>OPERATIONS OVERVIEW • MAY 2026</p>
            </div>
          </div>
          <div className="sync-status">
            <span className="dot"></span>
            LOCAL EXCEL DATA
          </div>
        </header>

        {/* KPI CARDS (Dynamically Populated Targets) */}
        <section className="kpi-grid">
          {[
            { label: "TOTAL HANDLED", val: summary.totalTickets, sub: `Target: ${summary.targetTickets}`, class: "blue-glow" },
            { label: "TOTAL SOLVED", val: summary.totalClosed, sub: `Target: ${summary.targetClosed}`, class: "green-glow" },
            { label: "L2 ESCALATION", val: `${summary.escalationRate}%`, sub: `Limit: ${summary.targetEscalation}`, class: "purple-glow" },
            { label: "CURRENT BACKLOG", val: summary.totalBacklog, sub: "Priority: High", class: "red-glow" },
          ].map((card, i) => (
            <div key={i} className={`glass-card kpi-card animate-pop delay-${i} ${card.class}`}>
              <p className="kpi-label">{card.label}</p>
              <h2 className="kpi-value">{card.val}</h2>
              <p className="kpi-sub">{card.sub}</p>
            </div>
          ))}
        </section>

        {/* ANALYTICS SECTION */}
        <section className="analytics-row">
          <div className="glass-card chart-box animate-pop delay-4">
            <h3>Weekly Target vs Actual</h3>
            <div className="chart-container">
              <Bar
                data={barChartData}
                options={{
                  responsive: true,
                  maintainAspectRatio: false,
                  plugins: { legend: { display: false } },
                }}
              />
            </div>
          </div>

          <div className="glass-card chart-box animate-pop delay-5">
            <h3>Resolution Efficiency</h3>
            <div className="doughnut-container">
              <Doughnut data={doughnutData} options={{ cutout: "80%" }} />
              <div className="doughnut-center">
                <p>Solved</p>
                <h4>{((summary.totalClosed / (summary.totalTickets || 1)) * 100).toFixed(0)}%</h4>
              </div>
            </div>
          </div>
        </section>

        {/* AGENT PERFORMANCE LEADERBOARD TABLE */}
        <section className="glass-card table-section animate-pop delay-6">
          <div className="table-header">
            <h3>Individual Performance Leaderboard</h3>
          </div>

          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Agent</th>
                  <th>New Tickets Owned</th>
                  <th>Escalated to L2</th>
                  <th>Solved/Closed</th>
                  <th>Unresolved Tickets</th>
                  <th>Efficiency</th>
                </tr>
              </thead>
              <tbody>
                {employees.map((emp, i) => (
                  <tr key={i}>
                    <td className="bold">{emp.name}</td>
                    <td>{emp.owned}</td>
                    <td>{emp.escalated}</td>
                    <td>{emp.closed}</td>
                    <td className="red-text">{emp.unresolved}</td>
                    <td>
                      <span className="efficiency-pill">
                        {emp.owned > 0 ? ((emp.closed / emp.owned) * 100).toFixed(0) : 0}%
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </div>
  );
}

export default App;