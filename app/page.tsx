"use client";

import { useEffect, useState, useMemo } from "react";

/* ─── CSV Export Utility ─── */
const escape = (v: unknown) => {
  const s = v == null ? "" : String(v);
  return s.includes(",") || s.includes('"') || s.includes("\n")
    ? `"${s.replace(/"/g, '""')}"` : s;
};

function rowsToCsvBlock(rows: Record<string, unknown>[]) {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  return [
    headers.join(","),
    ...rows.map((r) => headers.map((h) => escape(r[h])).join(",")),
  ].join("\n");
}

function exportCsv(rows: Record<string, unknown>[], filename: string) {
  if (!rows.length) return;
  const csv = rowsToCsvBlock(rows);
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function exportAllCsv(data: AppData, customerLinks: Record<string, string>, customerMap: Record<string, { code: string; name: string; phone: string; nickname: string }>) {
  const now = new Date().toISOString().slice(0, 10);

  // Build combined all-columns approach
  const allCols = [
    "ตาราง",
    // สาขา
    "รหัสสาขา","ชื่อสาขา",
    // ผู้ตรวจสอบ
    "รหัสผู้ตรวจ","ชื่อผู้ตรวจ","เบอร์ผู้ตรวจ","อีเมลผู้ตรวจ",
    // ลูกค้า
    "รหัสลูกค้า","ชื่อลูกค้า","ชื่อเล่น","เบอร์ลูกค้า","ที่อยู่ลูกค้า",
    // สัญญา
    "เลขสัญญา","สินค้า","ราคารวม","ผ่อนต่องวด","จำนวนงวด","วันเริ่ม","วันสิ้นสุด","สถานะสัญญา","พนักงานขาย","ยอดชำระแล้ว","งวดที่จ่ายแล้ว",
    // ชำระ
    "งวดที่","จำนวนเงิน","วันครบกำหนด","วันที่จ่าย","สถานะชำระ","หมายเหตุ",
  ];

  const empty = Object.fromEntries(allCols.map((c) => [c, ""]));

  const sections: Record<string, unknown>[][] = [];

  // Section: สาขา
  sections.push([{ ...empty, ตาราง: "=== สาขา ===" }]);
  sections.push(data.branches.map((r) => ({
    ...empty,
    ตาราง: "สาขา",
    รหัสสาขา: r["branch_code*"], ชื่อสาขา: r["branch_name*"],
  })));

  // Section: ผู้ตรวจสอบ
  sections.push([{ ...empty, ตาราง: "=== ผู้ตรวจสอบ ===" }]);
  sections.push(data.checkers.map((r) => ({
    ...empty,
    ตาราง: "ผู้ตรวจสอบ",
    รหัสสาขา: r["branch_code*"], ชื่อสาขา: "",
    รหัสผู้ตรวจ: r["checker_code*"],
    ชื่อผู้ตรวจ: String(r["name*"] ?? "") + " " + String(r["surname"] ?? ""),
    เบอร์ผู้ตรวจ: r["phone"], อีเมลผู้ตรวจ: r["email"],
  })));

  // Section: ลูกค้า
  sections.push([{ ...empty, ตาราง: "=== ลูกค้า ===" }]);
  sections.push(data.customers.map((r) => ({
    ...empty,
    ตาราง: "ลูกค้า",
    รหัสสาขา: r["branch_code*"],
    รหัสผู้ตรวจ: r["checker_code*"],
    รหัสลูกค้า: r["customer_code*"],
    ชื่อลูกค้า: String(r["name*"] ?? "") + " " + String(r["surname"] ?? ""),
    ชื่อเล่น: r["nickname"], เบอร์ลูกค้า: r["phone"], ที่อยู่ลูกค้า: r["address"],
  })));

  // Section: สัญญา
  sections.push([{ ...empty, ตาราง: "=== สัญญาผ่อนชำระ ===" }]);
  sections.push(data.installments.map((r) => {
    const cn = String(r.contract_number ?? "");
    const cust = customerLinks[cn] ? customerMap[customerLinks[cn]] : null;
    return {
      ...empty,
      ตาราง: "สัญญา",
      รหัสสาขา: r.branch_code, ชื่อสาขา: r.branch_name,
      เลขสัญญา: cn,
      รหัสลูกค้า: cust?.code ?? "",
      ชื่อลูกค้า: cust?.name ?? "",
      เบอร์ลูกค้า: cust?.phone ?? "",
      สินค้า: r.product_name,
      ราคารวม: r.total_amount, ผ่อนต่องวด: r.installment_amount,
      จำนวนงวด: r.installment_period, วันเริ่ม: r.start_date,
      วันสิ้นสุด: r.end_date, สถานะสัญญา: r.status,
      พนักงานขาย: r.salesperson_name,
      ยอดชำระแล้ว: r.total_paid, งวดที่จ่ายแล้ว: r.payment_count,
    };
  }));

  // Section: ประวัติชำระ
  sections.push([{ ...empty, ตาราง: "=== ประวัติการชำระ ===" }]);
  sections.push(data.payments.map((r) => ({
    ...empty,
    ตาราง: "ชำระ",
    เลขสัญญา: r.contract_number,
    งวดที่: r.payment_sequence, จำนวนเงิน: r.amount,
    วันครบกำหนด: r.due_date, วันที่จ่าย: r.payment_date,
    สถานะชำระ: r.status, หมายเหตุ: r.notes,
  })));

  const allRows = sections.flat();
  const csv = rowsToCsvBlock(allRows);
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `ข้อมูลทั้งหมด_${now}.csv`; a.click();
  URL.revokeObjectURL(url);

}

/* ─── Types ─── */
type Branch = Record<string, unknown>;
type Checker = Record<string, unknown>;
type Customer = Record<string, unknown>;
type Installment = {
  contract_number: unknown;
  product_name: unknown;
  total_amount: unknown;
  installment_amount: unknown;
  installment_period: unknown;
  start_date: unknown;
  end_date: unknown;
  status: unknown;
  branch_code: unknown;
  branch_name: string;
  salesperson_name: unknown;
  payment_count: number;
  total_paid: string;
};
type Payment = {
  contract_number: unknown;
  payment_sequence: unknown;
  amount: unknown;
  due_date: unknown;
  payment_date: unknown;
  status: unknown;
  notes: unknown;
};
type CombinedCustomer = {
  customer_code: unknown;
  name: string;
  nickname: unknown;
  phone: unknown;
  address: unknown;
  branch_code: unknown;
  branch_name: string;
  checker_code: unknown;
  checker_name: string;
};

type AppData = {
  branches: Branch[];
  checkers: Checker[];
  customers: Customer[];
  combined: CombinedCustomer[];
  installments: Installment[];
  payments: Payment[];
  paymentsByContract: Record<string, Payment[]>;
};

type Tab = "overview" | "branches" | "checkers" | "customers" | "installments" | "payments";

const TABS: { id: Tab; label: string; emoji: string }[] = [
  { id: "overview",     label: "ภาพรวม",         emoji: "🔗" },
  { id: "installments", label: "สัญญาผ่อนชำระ",  emoji: "📋" },
  { id: "payments",     label: "ประวัติชำระ",     emoji: "💳" },
  { id: "customers",    label: "ลูกค้า",           emoji: "👥" },
  { id: "checkers",     label: "ผู้ตรวจสอบ",      emoji: "👤" },
  { id: "branches",     label: "สาขา",             emoji: "🏢" },
];

const PAGE_SIZE = 50;

/* ─── Main ─── */
export default function Home() {
  const [data, setData] = useState<AppData | null>(null);
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState<Tab>("overview");
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [branchFilter, setBranchFilter] = useState("all");
  const [checkerFilter, setCheckerFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  // Payments drill-down
  const [selectedContract, setSelectedContract] = useState<string | null>(null);
  // contract_number → customer_code mapping (persisted in localStorage)
  const [customerLinks, setCustomerLinks] = useState<Record<string, string>>({});
  const [linkModal, setLinkModal] = useState<{ contractNumber: string } | null>(null);

  useEffect(() => {
    const saved = localStorage.getItem("customerLinks");
    if (saved) setCustomerLinks(JSON.parse(saved));
  }, []);

  function saveLink(contractNumber: string, customerCode: string) {
    setCustomerLinks((prev) => {
      const next = { ...prev, [contractNumber]: customerCode };
      localStorage.setItem("customerLinks", JSON.stringify(next));
      return next;
    });
  }

  function removeLink(contractNumber: string) {
    setCustomerLinks((prev) => {
      const next = { ...prev };
      delete next[contractNumber];
      localStorage.setItem("customerLinks", JSON.stringify(next));
      return next;
    });
  }

  useEffect(() => {
    fetch("/api/data")
      .then((r) => r.json())
      .then((d) => { setData(d); setLoading(false); });
  }, []);

  useEffect(() => { setPage(1); }, [activeTab, search, branchFilter, checkerFilter, statusFilter]);

  const branchOptions = useMemo(() => {
    if (!data) return [];
    return data.branches.map((b) => ({ code: String(b["branch_code*"]), name: String(b["branch_name*"]) }));
  }, [data]);

  const checkerOptions = useMemo(() => {
    if (!data) return [];
    return data.checkers.map((c) => ({
      code: String(c["checker_code*"]),
      name: `${c["name*"] ?? ""} ${c["surname"] ?? ""}`.trim(),
    }));
  }, [data]);

  const installmentStatuses = useMemo(() => {
    if (!data) return [];
    return [...new Set(data.installments.map((i) => String(i.status ?? "")))].filter(Boolean);
  }, [data]);

  // customer lookup: code → { name, phone, nickname }
  const customerMap = useMemo(() => {
    if (!data) return {} as Record<string, { code: string; name: string; phone: string; nickname: string }>;
    return Object.fromEntries(
      data.customers.map((c) => [
        String(c["customer_code*"]),
        {
          code: String(c["customer_code*"]),
          name: `${c["name*"] ?? ""} ${c["surname"] ?? ""}`.trim(),
          phone: String(c["phone"] ?? ""),
          nickname: String(c["nickname"] ?? ""),
        },
      ])
    );
  }, [data]);

  // ── Filtered data ──
  const filteredCombined = useMemo(() => {
    if (!data) return [];
    let rows = data.combined;
    if (branchFilter !== "all") rows = rows.filter((r) => String(r.branch_code) === branchFilter);
    if (checkerFilter !== "all") rows = rows.filter((r) => String(r.checker_code) === checkerFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      rows = rows.filter((r) =>
        String(r.customer_code).toLowerCase().includes(q) ||
        r.name.toLowerCase().includes(q) ||
        String(r.nickname ?? "").toLowerCase().includes(q) ||
        String(r.phone ?? "").toLowerCase().includes(q)
      );
    }
    return rows;
  }, [data, search, branchFilter, checkerFilter]);

  const filteredInstallments = useMemo(() => {
    if (!data) return [];
    let rows = data.installments;
    if (branchFilter !== "all") rows = rows.filter((r) => String(r.branch_code) === branchFilter);
    if (statusFilter !== "all") rows = rows.filter((r) => String(r.status) === statusFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      rows = rows.filter((r) =>
        String(r.contract_number ?? "").toLowerCase().includes(q) ||
        String(r.product_name ?? "").toLowerCase().includes(q) ||
        String(r.salesperson_name ?? "").toLowerCase().includes(q)
      );
    }
    return rows;
  }, [data, search, branchFilter, statusFilter]);

  const filteredPayments = useMemo(() => {
    if (!data) return [];
    let rows = data.payments;
    if (search.trim()) {
      const q = search.toLowerCase();
      rows = rows.filter((r) => String(r.contract_number ?? "").toLowerCase().includes(q));
    }
    return rows;
  }, [data, search]);

  const filteredCustomers = useMemo(() => {
    if (!data) return [];
    let rows = data.customers;
    if (branchFilter !== "all") rows = rows.filter((r) => String(r["branch_code*"]) === branchFilter);
    if (checkerFilter !== "all") rows = rows.filter((r) => String(r["checker_code*"]) === checkerFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      rows = rows.filter((r) =>
        String(r["customer_code*"] ?? "").toLowerCase().includes(q) ||
        String(r["name*"] ?? "").toLowerCase().includes(q) ||
        String(r["surname"] ?? "").toLowerCase().includes(q) ||
        String(r["phone"] ?? "").toLowerCase().includes(q)
      );
    }
    return rows;
  }, [data, search, branchFilter, checkerFilter]);

  function paginate<T>(arr: T[]) {
    return { items: arr.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE), total: arr.length };
  }

  // ── Summary stats ──
  const stats = useMemo(() => {
    if (!data) return null;
    const totalValue = data.installments.reduce((s, i) => s + parseFloat(String(i.total_amount ?? 0)), 0);
    const totalPaid = data.installments.reduce((s, i) => s + parseFloat(i.total_paid), 0);
    const activeContracts = data.installments.filter((i) => String(i.status).includes("ค้างชำระ") || String(i.status).includes("ปกติ")).length;
    return { totalValue, totalPaid, activeContracts };
  }, [data]);

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4" />
          <p className="text-gray-600 text-lg">กำลังโหลดข้อมูล...</p>
          <p className="text-gray-400 text-sm mt-1">กำลังอ่านไฟล์ Excel และ CSV 8 ไฟล์</p>
        </div>
      </div>
    );
  }

  if (!data) return null;

  const showBranchFilter = ["overview", "customers", "installments"].includes(activeTab);
  const showCheckerFilter = ["overview", "customers"].includes(activeTab);
  const showStatusFilter = activeTab === "installments";

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white border-b shadow-sm sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 py-3 flex flex-wrap items-center justify-between gap-3">
          <div>
            <h1 className="text-lg font-bold text-gray-800">📊 ระบบดูข้อมูล</h1>
          </div>
          <div className="flex flex-wrap gap-2 text-xs">
            <Chip color="blue">🏢 {data.branches.length} สาขา</Chip>
            <Chip color="purple">👤 {data.checkers.length} ผู้ตรวจสอบ</Chip>
            <Chip color="green">👥 {data.customers.length.toLocaleString()} ลูกค้า</Chip>
            <Chip color="orange">📋 {data.installments.length.toLocaleString()} สัญญา</Chip>
            <Chip color="red">💳 {data.payments.length.toLocaleString()} รายการชำระ</Chip>
          </div>
          <ExportAllButton data={data} customerLinks={customerLinks} customerMap={customerMap} />
        </div>

        {/* Tabs */}
        <div className="max-w-7xl mx-auto px-4 flex gap-1 overflow-x-auto">
          {TABS.map((t) => (
            <button key={t.id} onClick={() => { setActiveTab(t.id); setSearch(""); setSelectedContract(null); }}
              className={`px-4 py-2 text-sm font-medium rounded-t-lg whitespace-nowrap transition-colors ${
                activeTab === t.id ? "bg-blue-600 text-white" : "text-gray-600 hover:bg-gray-100"
              }`}>
              {t.emoji} {t.label}
            </button>
          ))}
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-5">

        {/* ── Stats bar (overview & installments) ── */}
        {(activeTab === "overview" || activeTab === "installments") && stats && (
          <div className="grid grid-cols-3 gap-4 mb-5">
            <StatCard label="มูลค่าสัญญารวม" value={`฿${stats.totalValue.toLocaleString("th-TH", { maximumFractionDigits: 0 })}`} color="blue" />
            <StatCard label="ยอดชำระสะสม" value={`฿${stats.totalPaid.toLocaleString("th-TH", { maximumFractionDigits: 0 })}`} color="green" />
            <StatCard label="สัญญาที่กำลังดำเนินการ" value={stats.activeContracts.toLocaleString()} color="orange" />
          </div>
        )}

        {/* ── Filters ── */}
        {(activeTab !== "branches" && activeTab !== "checkers") && (
          <div className="flex flex-wrap gap-2 mb-4">
            <input type="text" value={search} onChange={(e) => setSearch(e.target.value)}
              placeholder={
                activeTab === "payments" ? "ค้นหาเลขสัญญา..." :
                activeTab === "installments" ? "ค้นหาสัญญา / สินค้า / พนักงาน..." :
                "ค้นหา รหัส / ชื่อ / เบอร์โทร..."
              }
              className="flex-1 min-w-[200px] max-w-xs px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
            {showBranchFilter && (
              <select value={branchFilter} onChange={(e) => setBranchFilter(e.target.value)}
                className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                <option value="all">🏢 ทุกสาขา</option>
                {branchOptions.map((b) => <option key={b.code} value={b.code}>{b.name}</option>)}
              </select>
            )}
            {showCheckerFilter && (
              <select value={checkerFilter} onChange={(e) => setCheckerFilter(e.target.value)}
                className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                <option value="all">👤 ทุกผู้ตรวจสอบ</option>
                {checkerOptions.map((c) => <option key={c.code} value={c.code}>{c.name}</option>)}
              </select>
            )}
            {showStatusFilter && (
              <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}
                className="px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                <option value="all">📌 ทุกสถานะ</option>
                {installmentStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
            )}
            <ExportButton
              activeTab={activeTab}
              filteredCombined={filteredCombined}
              filteredInstallments={filteredInstallments}
              filteredPayments={filteredPayments}
              filteredCustomers={filteredCustomers}
              branches={data?.branches ?? []}
              checkers={data?.checkers ?? []}
              customerLinks={customerLinks}
              customerMap={customerMap}
            />
          </div>
        )}

        {/* ══════════ OVERVIEW TAB ══════════ */}
        {activeTab === "overview" && (
          <Section title="ภาพรวมลูกค้า (รวมสาขา + ผู้ตรวจสอบ)"
            count={filteredCombined.length} total={data.combined.length}>
            <table className="w-full text-sm border-collapse">
              <Thead cols={["รหัสลูกค้า","ชื่อ-นามสกุล","ชื่อเล่น","เบอร์โทร","สาขา","ผู้ตรวจสอบ","ที่อยู่"]} />
              <tbody>
                {paginate(filteredCombined).items.map((r, i) => (
                  <tr key={i} className="border-b hover:bg-blue-50 transition-colors">
                    <Td mono>{String(r.customer_code ?? "-")}</Td>
                    <Td bold>{r.name || "-"}</Td>
                    <Td muted>{String(r.nickname ?? "-")}</Td>
                    <Td>{String(r.phone ?? "-")}</Td>
                    <Td><Badge color="blue">{r.branch_name}</Badge></Td>
                    <Td><Badge color="purple">{r.checker_name}</Badge></Td>
                    <Td muted truncate>{String(r.address ?? "-")}</Td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Pagination page={page} setPage={setPage} total={filteredCombined.length} />
          </Section>
        )}

        {/* ══════════ INSTALLMENTS TAB ══════════ */}
        {activeTab === "installments" && (
          <>
          {linkModal && (
            <LinkCustomerModal
              contractNumber={linkModal.contractNumber}
              customers={data.customers}
              customerMap={customerMap}
              currentCode={customerLinks[linkModal.contractNumber]}
              onLink={(code) => { saveLink(linkModal.contractNumber, code); setLinkModal(null); }}
              onRemove={() => { removeLink(linkModal.contractNumber); setLinkModal(null); }}
              onClose={() => setLinkModal(null)}
            />
          )}
          <Section title="สัญญาผ่อนชำระ"
            count={filteredInstallments.length} total={data.installments.length}>
            <table className="w-full text-sm border-collapse">
              <Thead cols={["เลขสัญญา","ลูกค้า","สินค้า","ราคารวม","ผ่อน/งวด","งวด","เริ่มต้น","สถานะ","สาขา","ชำระแล้ว","งวดที่จ่าย","ประวัติ"]} />
              <tbody>
                {paginate(filteredInstallments).items.map((r, i) => {
                  const cn = String(r.contract_number ?? "");
                  const linkedCode = customerLinks[cn];
                  const linkedCust = linkedCode ? customerMap[linkedCode] : null;
                  const pct = r.total_amount
                    ? Math.min(100, (parseFloat(r.total_paid) / parseFloat(String(r.total_amount))) * 100)
                    : 0;
                  const isDone = String(r.status).includes("เสร็จสิ้น");
                  return (
                    <tr key={i} className="border-b hover:bg-orange-50 transition-colors">
                      <Td mono>{cn || "-"}</Td>
                      <td className="px-3 py-2 border-b whitespace-nowrap">
                        {linkedCust ? (
                          <div className="flex items-center gap-1.5">
                            <div>
                              <p className="font-medium text-gray-800 text-xs">{linkedCust.name}</p>
                              <p className="text-gray-400 text-xs">{linkedCust.phone}</p>
                            </div>
                            <button onClick={() => setLinkModal({ contractNumber: cn })}
                              className="text-gray-300 hover:text-blue-500 transition-colors ml-1" title="เปลี่ยน">✎</button>
                          </div>
                        ) : (
                          <button onClick={() => setLinkModal({ contractNumber: cn })}
                            className="flex items-center gap-1 text-xs text-blue-500 hover:text-blue-700 border border-dashed border-blue-300 hover:border-blue-500 rounded px-2 py-1 transition-colors">
                            <span>+</span> ลิงก์ลูกค้า
                          </button>
                        )}
                      </td>
                      <Td maxW>{String(r.product_name ?? "-")}</Td>
                      <Td right>฿{Number(r.total_amount).toLocaleString()}</Td>
                      <Td right>฿{Number(r.installment_amount).toLocaleString()}</Td>
                      <Td center>{String(r.installment_period ?? "-")}</Td>
                      <Td>{String(r.start_date ?? "-")}</Td>
                      <Td><Badge color={isDone ? "green" : "yellow"}>{String(r.status ?? "-")}</Badge></Td>
                      <Td><Badge color="blue">{r.branch_name}</Badge></Td>
                      <td className="px-3 py-2 border-b">
                        <div className="flex items-center gap-2">
                          <div className="flex-1 bg-gray-200 rounded-full h-1.5 min-w-[60px]">
                            <div className="bg-green-500 h-1.5 rounded-full" style={{ width: `${pct}%` }} />
                          </div>
                          <span className="text-xs text-gray-500 whitespace-nowrap">
                            ฿{Number(r.total_paid).toLocaleString()}
                          </span>
                        </div>
                      </td>
                      <Td center>{r.payment_count}</Td>
                      <td className="px-3 py-2 border-b">
                        <button
                          onClick={() => { setSelectedContract(cn); setActiveTab("payments"); setSearch(cn); }}
                          className="text-xs text-blue-600 hover:underline whitespace-nowrap">
                          ดู →
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <Pagination page={page} setPage={setPage} total={filteredInstallments.length} />
          </Section>
          </>
        )}

        {/* ══════════ PAYMENTS TAB ══════════ */}
        {activeTab === "payments" && (
          <>
            {selectedContract && (
              <div className="mb-3 flex items-center gap-2 text-sm">
                <span className="bg-blue-100 text-blue-700 px-3 py-1 rounded-full font-medium">
                  📋 กรองตามสัญญา: {selectedContract}
                </span>
                <button onClick={() => { setSelectedContract(null); setSearch(""); }}
                  className="text-gray-400 hover:text-red-500 text-xs">✕ ล้างตัวกรอง</button>
              </div>
            )}
            <Section title="ประวัติการชำระเงิน"
              count={filteredPayments.length} total={data.payments.length}>
              <table className="w-full text-sm border-collapse">
                <Thead cols={["เลขสัญญา","งวดที่","จำนวนเงิน","วันครบกำหนด","วันที่จ่าย","สถานะ","หมายเหตุ"]} />
                <tbody>
                  {paginate(filteredPayments).items.map((r, i) => (
                    <tr key={i} className="border-b hover:bg-green-50 transition-colors">
                      <td className="px-3 py-2">
                        <button onClick={() => { setSelectedContract(String(r.contract_number)); setSearch(String(r.contract_number)); }}
                          className="font-mono text-xs text-blue-600 hover:underline">
                          {String(r.contract_number ?? "-")}
                        </button>
                      </td>
                      <Td center>{String(r.payment_sequence ?? "-")}</Td>
                      <Td right>฿{Number(r.amount).toLocaleString()}</Td>
                      <Td>{String(r.due_date ?? "-")}</Td>
                      <Td>{String(r.payment_date ?? "-")}</Td>
                      <Td><Badge color="green">{String(r.status ?? "-")}</Badge></Td>
                      <Td muted>{String(r.notes ?? "-")}</Td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <Pagination page={page} setPage={setPage} total={filteredPayments.length} />
            </Section>
          </>
        )}

        {/* ══════════ CUSTOMERS TAB ══════════ */}
        {activeTab === "customers" && (
          <Section title="ข้อมูลลูกค้า"
            count={filteredCustomers.length} total={data.customers.length}>
            <table className="w-full text-sm border-collapse">
              <Thead cols={["รหัสลูกค้า","ชื่อ","นามสกุล","ชื่อเล่น","เบอร์โทร","สาขา","ผู้ตรวจสอบ"]} />
              <tbody>
                {paginate(filteredCustomers).items.map((r, i) => (
                  <tr key={i} className="border-b hover:bg-gray-50 transition-colors">
                    <Td mono>{String(r["customer_code*"] ?? "-")}</Td>
                    <Td bold>{String(r["name*"] ?? "-")}</Td>
                    <Td>{String(r["surname"] ?? "-")}</Td>
                    <Td muted>{String(r["nickname"] ?? "-")}</Td>
                    <Td>{String(r["phone"] ?? "-")}</Td>
                    <Td><Badge color="blue">{String(r["branch_code*"] ?? "-")}</Badge></Td>
                    <Td><Badge color="purple">{String(r["checker_code*"] ?? "-")}</Badge></Td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Pagination page={page} setPage={setPage} total={filteredCustomers.length} />
          </Section>
        )}

        {/* ══════════ CHECKERS TAB ══════════ */}
        {activeTab === "checkers" && (
          <>
          <div className="flex justify-end mb-3">
            <ExportButton activeTab="checkers" filteredCombined={[]} filteredInstallments={[]} filteredPayments={[]} filteredCustomers={[]} branches={data.branches} checkers={data.checkers} customerLinks={{}} customerMap={{}} />
          </div>
          <Section title="ข้อมูลผู้ตรวจสอบ" count={data.checkers.length}>
            <table className="w-full text-sm border-collapse">
              <Thead cols={["รหัส","ชื่อ","นามสกุล","เบอร์โทร","อีเมล","รหัสสาขา"]} />
              <tbody>
                {data.checkers.map((r, i) => (
                  <tr key={i} className="border-b hover:bg-gray-50 transition-colors">
                    <Td mono>{String(r["checker_code*"] ?? "-")}</Td>
                    <Td bold>{String(r["name*"] ?? "-")}</Td>
                    <Td>{String(r["surname"] ?? "-")}</Td>
                    <Td>{String(r["phone"] ?? "-")}</Td>
                    <Td>{String(r["email"] ?? "-")}</Td>
                    <Td><Badge color="blue">{String(r["branch_code*"] ?? "-")}</Badge></Td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Section>
          </>
        )}

        {/* ══════════ BRANCHES TAB ══════════ */}
        {activeTab === "branches" && (
          <>
          <div className="flex justify-end mb-3">
            <ExportButton activeTab="branches" filteredCombined={[]} filteredInstallments={[]} filteredPayments={[]} filteredCustomers={[]} branches={data.branches} checkers={data.checkers} customerLinks={{}} customerMap={{}} />
          </div>
          <Section title="ข้อมูลสาขา" count={data.branches.length}>
            <table className="w-full text-sm border-collapse">
              <Thead cols={["รหัสสาขา","ชื่อสาขา","ที่อยู่","เบอร์โทร","ผู้จัดการ"]} />
              <tbody>
                {data.branches.map((r, i) => (
                  <tr key={i} className="border-b hover:bg-gray-50 transition-colors">
                    <Td mono>{String(r["branch_code*"] ?? "-")}</Td>
                    <Td bold>{String(r["branch_name*"] ?? "-")}</Td>
                    <Td>{String(r["address"] ?? "-")}</Td>
                    <Td>{String(r["phone"] ?? "-")}</Td>
                    <Td>{String(r["manager"] ?? "-")}</Td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Section>
          </>
        )}

      </main>
    </div>
  );
}

/* ─── Small Components ─── */

function Chip({ children, color }: { children: React.ReactNode; color: string }) {
  const colors: Record<string, string> = {
    blue: "bg-blue-100 text-blue-700",
    purple: "bg-purple-100 text-purple-700",
    green: "bg-green-100 text-green-700",
    orange: "bg-orange-100 text-orange-700",
    red: "bg-red-100 text-red-700",
  };
  return <span className={`px-2.5 py-1 rounded-full font-medium ${colors[color]}`}>{children}</span>;
}

function StatCard({ label, value, color }: { label: string; value: string; color: string }) {
  const colors: Record<string, string> = {
    blue: "border-blue-200 bg-blue-50 text-blue-700",
    green: "border-green-200 bg-green-50 text-green-700",
    orange: "border-orange-200 bg-orange-50 text-orange-700",
  };
  return (
    <div className={`rounded-xl border p-4 ${colors[color]}`}>
      <p className="text-xs font-medium opacity-70 mb-1">{label}</p>
      <p className="text-2xl font-bold">{value}</p>
    </div>
  );
}

function Section({ title, count, total, children }: { title: string; count: number; total?: number; children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-xl shadow-sm border overflow-hidden">
      <div className="px-5 py-3 border-b flex items-center justify-between bg-gray-50">
        <h2 className="font-semibold text-gray-800">{title}</h2>
        <span className="text-sm text-gray-500">
          {total && count !== total
            ? `แสดง ${count.toLocaleString()} จาก ${total.toLocaleString()} รายการ`
            : `${count.toLocaleString()} รายการ`}
        </span>
      </div>
      <div className="overflow-auto">{children}</div>
    </div>
  );
}

function Thead({ cols }: { cols: string[] }) {
  return (
    <thead>
      <tr className="bg-gray-100 text-gray-700">
        {cols.map((c) => (
          <th key={c} className="px-3 py-2.5 text-left font-semibold border-b whitespace-nowrap text-sm">{c}</th>
        ))}
      </tr>
    </thead>
  );
}

function Td({ children, mono, bold, muted, right, center, truncate, maxW }: {
  children: React.ReactNode; mono?: boolean; bold?: boolean; muted?: boolean;
  right?: boolean; center?: boolean; truncate?: boolean; maxW?: boolean;
}) {
  return (
    <td className={`px-3 py-2 border-b ${mono ? "font-mono text-xs text-gray-500" : ""} ${bold ? "font-medium" : ""} ${muted ? "text-gray-400" : "text-gray-700"} ${right ? "text-right" : ""} ${center ? "text-center" : ""} ${truncate ? "max-w-[200px] truncate" : ""} ${maxW ? "max-w-[220px] truncate" : ""} whitespace-nowrap`}>
      {children === null || children === undefined || children === "null" || children === "undefined" ? (
        <span className="text-gray-300">-</span>
      ) : children}
    </td>
  );
}

function Badge({ children, color }: { children: React.ReactNode; color: string }) {
  const colors: Record<string, string> = {
    blue: "bg-blue-100 text-blue-700",
    purple: "bg-purple-100 text-purple-700",
    green: "bg-green-100 text-green-700",
    yellow: "bg-yellow-100 text-yellow-700",
    red: "bg-red-100 text-red-600",
    orange: "bg-orange-100 text-orange-700",
  };
  return (
    <span className={`px-2 py-0.5 rounded-full text-xs font-medium whitespace-nowrap ${colors[color] ?? "bg-gray-100 text-gray-600"}`}>
      {children}
    </span>
  );
}

function Pagination({ page, setPage, total }: { page: number; setPage: (p: number) => void; total: number }) {
  const totalPages = Math.ceil(total / PAGE_SIZE);
  if (totalPages <= 1) return null;
  const start = (page - 1) * PAGE_SIZE + 1;
  const end = Math.min(page * PAGE_SIZE, total);
  return (
    <div className="flex items-center justify-between px-5 py-3 border-t bg-gray-50 text-sm text-gray-600">
      <span>แสดง {start.toLocaleString()}–{end.toLocaleString()} จาก {total.toLocaleString()} รายการ</span>
      <div className="flex gap-1">
        <PageBtn onClick={() => setPage(1)} disabled={page === 1}>«</PageBtn>
        <PageBtn onClick={() => setPage(page - 1)} disabled={page === 1}>ก่อนหน้า</PageBtn>
        <span className="px-3 py-1 rounded border bg-blue-600 text-white font-medium">{page} / {totalPages}</span>
        <PageBtn onClick={() => setPage(page + 1)} disabled={page === totalPages}>ถัดไป</PageBtn>
        <PageBtn onClick={() => setPage(totalPages)} disabled={page === totalPages}>»</PageBtn>
      </div>
    </div>
  );
}

function PageBtn({ children, onClick, disabled }: { children: React.ReactNode; onClick: () => void; disabled: boolean }) {
  return (
    <button onClick={onClick} disabled={disabled}
      className="px-3 py-1 rounded border disabled:opacity-40 hover:bg-gray-200 transition-colors">
      {children}
    </button>
  );
}

/* ─── Export Button ─── */
type CustomerInfo = { code: string; name: string; phone: string; nickname: string };

type ExportButtonProps = {
  activeTab: Tab;
  filteredCombined: CombinedCustomer[];
  filteredInstallments: Installment[];
  filteredPayments: Payment[];
  filteredCustomers: Customer[];
  branches: Branch[];
  checkers: Checker[];
  customerLinks: Record<string, string>;
  customerMap: Record<string, CustomerInfo>;
};

function ExportAllButton({ data, customerLinks, customerMap }: { data: AppData; customerLinks: Record<string, string>; customerMap: Record<string, CustomerInfo> }) {
  const [exporting, setExporting] = useState(false);
  return (
    <button
      onClick={() => { setExporting(true); setTimeout(() => { exportAllCsv(data, customerLinks, customerMap); setExporting(false); }, 10); }}
      disabled={exporting}
      className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-sm font-semibold rounded-lg transition-colors shadow-sm whitespace-nowrap"
    >
      {exporting ? <span className="animate-spin inline-block">⏳</span> : <span>⬇️</span>}
      Export ทั้งหมด (.csv)
    </button>
  );
}

function ExportButton({ activeTab, filteredCombined, filteredInstallments, filteredPayments, filteredCustomers, branches, checkers, customerLinks, customerMap }: ExportButtonProps) {
  const [exporting, setExporting] = useState(false);

  const handleExport = () => {
    setExporting(true);
    setTimeout(() => {
      const now = new Date().toISOString().slice(0, 10);
      if (activeTab === "overview") {
        exportCsv(
          filteredCombined.map((r) => ({
            รหัสลูกค้า: r.customer_code,
            ชื่อนามสกุล: r.name,
            ชื่อเล่น: r.nickname,
            เบอร์โทร: r.phone,
            ที่อยู่: r.address,
            รหัสสาขา: r.branch_code,
            ชื่อสาขา: r.branch_name,
            รหัสผู้ตรวจสอบ: r.checker_code,
            ชื่อผู้ตรวจสอบ: r.checker_name,
          })),
          `overview_${now}.csv`
        );
      } else if (activeTab === "installments") {
        exportCsv(
          filteredInstallments.map((r) => {
            const cn = String(r.contract_number ?? "");
            const cust = customerLinks[cn] ? customerMap[customerLinks[cn]] : null;
            return {
              เลขสัญญา: cn,
              รหัสลูกค้า: cust?.code ?? "",
              ชื่อลูกค้า: cust?.name ?? "",
              เบอร์ลูกค้า: cust?.phone ?? "",
              สินค้า: r.product_name,
              ราคารวม: r.total_amount,
              ผ่อนต่องวด: r.installment_amount,
              จำนวนงวด: r.installment_period,
              วันเริ่ม: r.start_date,
              วันสิ้นสุด: r.end_date,
              สถานะ: r.status,
              สาขา: r.branch_name,
              พนักงานขาย: r.salesperson_name,
              ยอดชำระแล้ว: r.total_paid,
              จำนวนงวดที่จ่าย: r.payment_count,
            };
          }),
          `installments_${now}.csv`
        );
      } else if (activeTab === "payments") {
        exportCsv(
          filteredPayments.map((r) => ({
            เลขสัญญา: r.contract_number,
            งวดที่: r.payment_sequence,
            จำนวนเงิน: r.amount,
            วันครบกำหนด: r.due_date,
            วันที่จ่าย: r.payment_date,
            สถานะ: r.status,
            หมายเหตุ: r.notes,
          })),
          `payments_${now}.csv`
        );
      } else if (activeTab === "customers") {
        exportCsv(
          filteredCustomers.map((r) => ({
            รหัสลูกค้า: r["customer_code*"],
            ชื่อ: r["name*"],
            นามสกุล: r["surname"],
            ชื่อเล่น: r["nickname"],
            เลขบัตรประชาชน: r["id_card*"],
            เบอร์โทร: r["phone"],
            อีเมล: r["email"],
            ที่อยู่: r["address"],
            ผู้ค้ำประกัน: r["guarantor_name"],
            เบอร์ผู้ค้ำ: r["guarantor_phone"],
            รหัสสาขา: r["branch_code*"],
            รหัสผู้ตรวจสอบ: r["checker_code*"],
          })),
          `customers_${now}.csv`
        );
      } else if (activeTab === "checkers") {
        exportCsv(
          checkers.map((r) => ({
            รหัส: r["checker_code*"],
            ชื่อ: r["name*"],
            นามสกุล: r["surname"],
            เบอร์โทร: r["phone"],
            อีเมล: r["email"],
            รหัสสาขา: r["branch_code*"],
          })),
          `checkers_${now}.csv`
        );
      } else if (activeTab === "branches") {
        exportCsv(
          branches.map((r) => ({
            รหัสสาขา: r["branch_code*"],
            ชื่อสาขา: r["branch_name*"],
            ที่อยู่: r["address"],
            เบอร์โทร: r["phone"],
            ผู้จัดการ: r["manager"],
          })),
          `branches_${now}.csv`
        );
      }
      setExporting(false);
    }, 10);
  };

  return (
    <button
      onClick={handleExport}
      disabled={exporting}
      className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-sm font-medium rounded-lg transition-colors shadow-sm"
    >
      {exporting ? <span className="animate-spin">⏳</span> : <span>⬇️</span>}
      Export CSV
    </button>
  );
}

/* ─── Link Customer Modal ─── */
function LinkCustomerModal({
  contractNumber,
  customers,
  customerMap,
  currentCode,
  onLink,
  onRemove,
  onClose,
}: {
  contractNumber: string;
  customers: Customer[];
  customerMap: Record<string, CustomerInfo>;
  currentCode?: string;
  onLink: (code: string) => void;
  onRemove: () => void;
  onClose: () => void;
}) {
  const [q, setQ] = useState("");

  const filtered = useMemo(() => {
    if (!q.trim()) return customers.slice(0, 50);
    const lower = q.toLowerCase();
    return customers
      .filter((c) =>
        String(c["customer_code*"] ?? "").toLowerCase().includes(lower) ||
        String(c["name*"] ?? "").toLowerCase().includes(lower) ||
        String(c["surname"] ?? "").toLowerCase().includes(lower) ||
        String(c["phone"] ?? "").toLowerCase().includes(lower) ||
        String(c["nickname"] ?? "").toLowerCase().includes(lower)
      )
      .slice(0, 50);
  }, [q, customers]);

  const current = currentCode ? customerMap[currentCode] : null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg mx-4 flex flex-col max-h-[80vh]">
        {/* Header */}
        <div className="px-5 py-4 border-b flex items-start justify-between">
          <div>
            <h2 className="font-bold text-gray-800">🔗 ลิงก์ลูกค้าให้สัญญา</h2>
            <p className="text-sm text-gray-500 mt-0.5">สัญญา: <span className="font-mono font-semibold text-blue-600">{contractNumber}</span></p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-xl leading-none">✕</button>
        </div>

        {/* Current link */}
        {current && (
          <div className="mx-5 mt-3 p-3 bg-blue-50 border border-blue-200 rounded-lg flex items-center justify-between">
            <div>
              <p className="text-xs text-blue-500 font-medium">ลูกค้าปัจจุบัน</p>
              <p className="font-semibold text-gray-800">{current.name}</p>
              <p className="text-xs text-gray-500">{current.code} · {current.phone}</p>
            </div>
            <button onClick={onRemove}
              className="text-xs text-red-500 hover:text-red-700 border border-red-200 hover:border-red-400 rounded px-2 py-1 transition-colors">
              ลบลิงก์
            </button>
          </div>
        )}

        {/* Search */}
        <div className="px-5 pt-3">
          <input
            autoFocus
            type="text" value={q} onChange={(e) => setQ(e.target.value)}
            placeholder="ค้นหา รหัส / ชื่อ / เบอร์โทร / ชื่อเล่น..."
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <p className="text-xs text-gray-400 mt-1">แสดง {filtered.length} รายการ{!q && " (พิมพ์เพื่อค้นหา)"}</p>
        </div>

        {/* List */}
        <div className="overflow-y-auto flex-1 px-5 pb-4 mt-2 space-y-1">
          {filtered.map((c) => {
            const code = String(c["customer_code*"]);
            const name = `${c["name*"] ?? ""} ${c["surname"] ?? ""}`.trim();
            const phone = String(c["phone"] ?? "");
            const nick = String(c["nickname"] ?? "");
            const isSelected = code === currentCode;
            return (
              <button key={code} onClick={() => onLink(code)}
                className={`w-full text-left px-3 py-2.5 rounded-lg transition-colors flex items-center justify-between gap-2 ${
                  isSelected ? "bg-blue-600 text-white" : "hover:bg-gray-100 text-gray-800"
                }`}>
                <div>
                  <span className="font-medium">{name}</span>
                  {nick && nick !== "null" && <span className={`ml-1.5 text-xs ${isSelected ? "text-blue-200" : "text-gray-400"}`}>({nick})</span>}
                  <p className={`text-xs mt-0.5 ${isSelected ? "text-blue-200" : "text-gray-400"}`}>{code} · {phone}</p>
                </div>
                {isSelected && <span className="text-white text-sm">✓</span>}
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}
