import { NextResponse } from "next/server";
import * as XLSX from "xlsx";
import path from "path";
import fs from "fs";

function readExcel(filename: string) {
  const filePath = path.join(process.cwd(), "public", "data", filename);
  const buffer = fs.readFileSync(filePath);
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: null }) as Record<string, unknown>[];
}

function readCsv(filename: string) {
  const filePath = path.join(process.cwd(), "public", "data", filename);
  const buffer = fs.readFileSync(filePath);
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: null }) as Record<string, unknown>[];
}

// Map สาย1/สาย2/สาย4 → branch_code 1/2/4
const BRANCH_CODE_MAP: Record<string, number> = {
  "สาย1": 1,
  "สาย2": 2,
  "สาย4": 4,
};

export async function GET() {
  const branches = readExcel("Branches.xlsx");
  const checkers = readExcel("Checkers.xlsx");
  const customers = readExcel("Customers.xlsx");
  const installmentsRaw = readExcel("Installments.xlsx");

  // Read all 6 payment CSV files and merge
  const payments: Record<string, unknown>[] = [];
  for (let i = 1; i <= 6; i++) {
    const rows = readCsv(`Payments-${i}.csv`);
    payments.push(...rows);
  }

  // Build lookup maps
  const branchMap = Object.fromEntries(
    branches.map((b) => [b["branch_code*"], b["branch_name*"]])
  );
  const checkerMap = Object.fromEntries(
    checkers.map((c) => [
      c["checker_code*"],
      `${c["name*"] ?? ""} ${c["surname"] ?? ""}`.trim(),
    ])
  );

  // Installments: contract_number* is empty; customer_code* holds the F-code (= contract number)
  // Map branch: "สาย1" → 1, etc.
  const installments = installmentsRaw.map((row) => ({
    contract_number: row["customer_code*"],          // F-code is actually contract number
    product_name: row["product_name*"],
    total_amount: row["total_amount*"],
    installment_amount: row["installment_amount*"],
    installment_period: row["installment_period*"],
    start_date: row["start_date*"],
    end_date: row["end_date*"],
    status: row["status*"],
    branch_code_raw: row["branch_code*"],
    branch_code: BRANCH_CODE_MAP[row["branch_code*"] as string] ?? null,
    branch_name: branchMap[BRANCH_CODE_MAP[row["branch_code*"] as string]] ?? String(row["branch_code*"] ?? "-"),
    salesperson_name: row["salesperson_name"],
  }));

  // Build payment lookup per contract
  const paymentsByContract: Record<string, Record<string, unknown>[]> = {};
  for (const p of payments) {
    const cn = String(p["contract_number"]);
    if (!paymentsByContract[cn]) paymentsByContract[cn] = [];
    paymentsByContract[cn].push(p);
  }

  // Customers combined view
  const combined = customers.map((c) => ({
    customer_code: c["customer_code*"],
    name: `${c["name*"] ?? ""} ${c["surname"] ?? ""}`.trim(),
    nickname: c["nickname"],
    phone: c["phone"],
    address: c["address"],
    branch_code: c["branch_code*"],
    branch_name: branchMap[c["branch_code*"] as number] ?? "-",
    checker_code: c["checker_code*"],
    checker_name: checkerMap[c["checker_code*"] as number] ?? "-",
  }));

  // Installments enriched with payment summary
  const installmentsWithSummary = installments.map((inst) => {
    const cn = String(inst.contract_number);
    const pays = paymentsByContract[cn] ?? [];
    const totalPaid = pays.reduce((s, p) => s + parseFloat(String(p["amount"] ?? 0)), 0);
    return {
      ...inst,
      payment_count: pays.length,
      total_paid: totalPaid.toFixed(2),
    };
  });

  return NextResponse.json({
    branches,
    checkers,
    customers,
    combined,
    installments: installmentsWithSummary,
    payments,
    paymentsByContract,
  });
}
