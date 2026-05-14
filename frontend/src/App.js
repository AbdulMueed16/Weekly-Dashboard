import React, { useEffect, useState } from "react";
import axios from "axios";

import {
  Bar
} from "react-chartjs-2";

import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Tooltip,
  Legend
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
  const [summary, setSummary] = useState({});

  const fetchDashboardData = () => {

    console.log("Fetching latest data...");

    axios.get("/api/dashboard")
      .then((response) => {

        console.log(response.data);

        setEmployees(response.data.employees);
        setSummary(response.data.summary);

      })
      .catch((error) => {

        console.log(error);

      });
  };

  useEffect(() => {

    fetchDashboardData();

    const interval = setInterval(() => {
      fetchDashboardData();
    }, 5000);

    return () => clearInterval(interval);

  }, []);

    const chartData = {
    labels: employees.map(emp => emp.name),
    datasets: [
      {
        label: "Closed Tickets",
        data: employees.map(emp => emp["total closed"]),
        backgroundColor: "#2563eb",
        borderRadius: 6
      },
      {
        label: "Weekly Targets",
        data: employees.map(emp => emp["weekly targets"]),
        backgroundColor: "#10b981",
        borderRadius: 6
      }
    ]
  };

  return (

    <div className="container">

      <h1>Weekly Team Dashboard</h1>

      <div className="cards">

        <div className="card">
          <h3>Total Employees</h3>
          <p>{summary.totalEmployees}</p>
        </div>

        <div className="card">
          <h3>Total Tickets</h3>
          <p>{summary.totalTickets}</p>
        </div>

        <div className="card">
          <h3>Total Closed</h3>
          <p>{summary.totalClosed}</p>
        </div>

        <div className="card">
          <h3>Total Backlog</h3>
          <p>{summary.totalBacklog}</p>
        </div>

      </div>

      <div className="chart-container">

        {employees.length > 0 && (
          <Bar data={chartData} />
        )}

      </div>

      <table>

        <thead>
          <tr>
            <th>Name</th>
            <th>Total Tickets</th>
            <th>Closed</th>
            <th>Moved</th>
            <th>Backlog</th>
            <th>Weekly Target</th>
          </tr>
        </thead>

        <tbody>

          {employees.map((emp, index) => (

            <tr key={index}>
              <td>{emp.name}</td>
              <td>{emp["total new tickets"]}</td>
              <td>{emp["total closed"]}</td>
              <td>{emp.moved}</td>
              <td>{emp.backlog}</td>
              <td>{emp["weekly targets"]}</td>
            </tr>

          ))}

        </tbody>

      </table>

    </div>
  );
}

export default App;