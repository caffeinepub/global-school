import type { PaymentReportRow } from "./reports";

export async function exportToExcel(
  rows: PaymentReportRow[],
  filename: string,
): Promise<void> {
  const { utils, writeFile } = await import("xlsx");

  const data = [
    ["Receipt No", "Parent Name", "Class", "Student Name", "Amount Paid (INR)"],
    ...rows.map((row) => [
      row.receiptNo,
      row.parentName,
      row.classLabel,
      row.studentName,
      row.amountPaid,
    ]),
  ];

  const worksheet = utils.aoa_to_sheet(data);

  // Set column widths
  worksheet["!cols"] = [
    { wch: 12 },
    { wch: 25 },
    { wch: 15 },
    { wch: 25 },
    { wch: 20 },
  ];

  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, worksheet, "Payments");

  writeFile(workbook, `${filename}.xlsx`);
}
