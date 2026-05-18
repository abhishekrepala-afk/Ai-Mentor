import { useEffect, useState } from "react";
import { callApi } from "../utils/api";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

function EnrollmentsPage() {
  const [enrollments, setEnrollments] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Dropdown state
  const [showDropdown, setShowDropdown] = useState(false);

  useEffect(() => {
    const fetchEnrollments = async () => {
      try {
        setLoading(true);

        const response = await callApi("/admin/enrollments?type=list");

        const enrollmentsList = Array.isArray(response?.data)
          ? response.data
          : Array.isArray(response)
          ? response
          : [];

        setEnrollments(enrollmentsList);
      } catch (err) {
        setError(err.message);
      } finally {
        setLoading(false);
      }
    };

    fetchEnrollments();
  }, []);

  // =========================
  // Excel Export (ExcelJS)
  // =========================
  const exportToExcel = async () => {
    try {
      const excelData = enrollments.map((enrollment) => ({
        Student: enrollment.user || "N/A",
        Email: enrollment.email || "N/A",
        Course: enrollment.course || "N/A",
        Date: enrollment.date
          ? new Date(enrollment.date).toLocaleDateString()
          : "N/A",
        Amount: `Rs ${enrollment.amount ?? 0}`,
        Status: enrollment.status || "Completed",
      }));

      // Create workbook and worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Enrollments");

      // Define columns
      worksheet.columns = [
        { header: "Student", key: "Student", width: 25 },
        { header: "Email", key: "Email", width: 30 },
        { header: "Course", key: "Course", width: 30 },
        { header: "Date", key: "Date", width: 15 },
        { header: "Amount", key: "Amount", width: 15 },
        { header: "Status", key: "Status", width: 20 },
      ];

      // Add rows
      excelData.forEach((row) => {
        worksheet.addRow(row);
      });

      // Style header row
      const headerRow = worksheet.getRow(1);
      headerRow.font = {
        bold: true,
        color: { argb: "FFFFFFFF" },
      };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "4472C4" },
      };
      headerRow.alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      // Center all cells
      worksheet.eachRow((row) => {
        row.eachCell((cell) => {
          cell.alignment = {
            vertical: "middle",
            horizontal: "center",
          };
        });
      });

      // Generate Excel buffer
      const buffer = await workbook.xlsx.writeBuffer();

      // Create blob
      const blob = new Blob([buffer], {
        type:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      // Download file
      saveAs(blob, "Enrollments_Report.xlsx");
    } catch (err) {
      console.error("Error exporting Excel:", err);
      alert("Failed to export Excel file.");
    }
  };

  // =========================
  // PDF Export
  // =========================
  const exportToPDF = () => {
    const doc = new jsPDF();

    const tableColumn = [
      "Student",
      "Email",
      "Course",
      "Date",
      "Amount",
      "Status",
    ];

    const tableRows = enrollments.map((enrollment) => [
      enrollment.user || "N/A",
      enrollment.email || "N/A",
      enrollment.course || "N/A",
      enrollment.date
        ? new Date(enrollment.date).toLocaleDateString()
        : "N/A",
      `Rs ${enrollment.amount ?? 0}`,
      enrollment.status || "Completed",
    ]);

    doc.setFontSize(16);
    doc.text("Enrollments Report", 14, 15);

    autoTable(doc, {
      head: [tableColumn],
      body: tableRows,
      startY: 25,
      styles: {
        fontSize: 9,
        halign: "center",
      },
      headStyles: {
        fillColor: [68, 114, 196],
      },
    });

    doc.save("Enrollments_Report.pdf");
  };

  if (loading) {
    return (
      <div className="p-10 text-center text-muted">
        Loading enrollments...
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-10 text-center text-red-500">
        Error: {error}
      </div>
    );
  }

  return (
    <>
      <div className="border-b border-border p-6 md:p-8 flex items-center justify-between">
        <h2 className="text-3xl font-semibold">
          All Enrollments
        </h2>

        {/* Export Dropdown */}
        <div className="relative">
          <button
            type="button"
            onClick={() => setShowDropdown(!showDropdown)}
            className="h-10 px-4 rounded-xl border border-border hover:bg-canvas-alt transition-colors"
          >
            Export Report
          </button>

          {showDropdown && (
            <div className="absolute right-0 mt-2 w-48 bg-canvas border border-border rounded-xl shadow-lg z-10 overflow-hidden">
              <button
                onClick={() => {
                  exportToExcel();
                  setShowDropdown(false);
                }}
                className="block w-full text-left px-4 py-3 hover:bg-canvas-alt transition-colors"
              >
                Download Excel
              </button>

              <button
                onClick={() => {
                  exportToPDF();
                  setShowDropdown(false);
                }}
                className="block w-full text-left px-4 py-3 hover:bg-canvas-alt transition-colors"
              >
                Download PDF
              </button>
            </div>
          )}
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full min-w-[900px]">
          <thead className="text-left text-xs uppercase tracking-wider text-muted">
            <tr className="border-b border-border">
              <th className="p-5">Student</th>
              <th>Email</th>
              <th>Course</th>
              <th>Date</th>
              <th>Amount</th>
              <th>Status</th>
            </tr>
          </thead>

          <tbody className="text-sm">
            {enrollments.length > 0 ? (
              enrollments.map((enrollment, index) => (
                <tr
                  key={`${enrollment.user}-${enrollment.course}-${index}`}
                  className="border-b border-border hover:bg-canvas-alt transition-colors"
                >
                  <td className="p-5">
                    <div className="font-medium">
                      {enrollment.user}
                    </div>
                  </td>

                  <td>{enrollment.email}</td>

                  <td>{enrollment.course}</td>

                  <td>
                    {enrollment.date
                      ? new Date(
                          enrollment.date
                        ).toLocaleDateString()
                      : "N/A"}
                  </td>

                  <td className="font-semibold">
                    Rs {enrollment.amount}
                  </td>

                  <td className="text-green-600">
                    {enrollment.status || "Completed"}
                  </td>
                </tr>
              ))
            ) : (
              <tr>
                <td
                  colSpan="6"
                  className="p-10 text-center text-muted italic"
                >
                  No enrollments found.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </>
  );
}

export default EnrollmentsPage;