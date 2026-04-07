import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  Home,
  MapPin,
  Calculator,
  DollarSign,
  TrendingUp,
  TrendingDown,
  Pencil,
  PlusCircle,
  Save,
  Trash2,
  Printer,
  Eye,
  EyeOff,
  RefreshCw,
  Upload,
} from "lucide-react";

const STORAGE_KEY = "mh_pricing_dataset_v1";
const GOOGLE_HOMES_URL_KEY = "mh_google_homes_url";
const GOOGLE_DEFAULTS_URL_KEY = "mh_google_defaults_url";
const LOGO_SRC = "/logo.png";

type UnitRow = {
  unit: string;
  currentRent: number;
  monthlyRent: number;
  lotRent: number;
  cost: number;
  terms: Record<string, { totalPI: number; taxesInsurance: number; profit?: number }>;
};

type ParkRow = {
  park: string;
  units: UnitRow[];
};

const REAL_PARK_DATA: ParkRow[] = [
  {
    park: "Breeze Pointe MHP",
    units: [
      {
        unit: "3",
        currentRent: 1100,
        monthlyRent: 1100,
        lotRent: 350,
        cost: 20000,
        terms: {
          "7": { totalPI: 63000, taxesInsurance: 12600 },
          "9": { totalPI: 81000, taxesInsurance: 16200 },
          "11": { totalPI: 99000, taxesInsurance: 19800 },
          "13": { totalPI: 117000, taxesInsurance: 23400 },
        },
      },
      {
        unit: "4",
        currentRent: 1250,
        monthlyRent: 1250,
        lotRent: 350,
        cost: 20000,
        terms: {
          "7": { totalPI: 75600, taxesInsurance: 12600 },
          "9": { totalPI: 97200, taxesInsurance: 16200 },
          "11": { totalPI: 118800, taxesInsurance: 19800 },
          "13": { totalPI: 140400, taxesInsurance: 23400 },
        },
      },
      {
        unit: "6",
        currentRent: 895,
        monthlyRent: 910,
        lotRent: 350,
        cost: 20000,
        terms: {
          "7": { totalPI: 47040, taxesInsurance: 12600 },
          "9": { totalPI: 60480, taxesInsurance: 16200 },
          "11": { totalPI: 73920, taxesInsurance: 19800 },
          "13": { totalPI: 87360, taxesInsurance: 23400 },
        },
      },
    ],
  },
  {
    park: "Fox Tail Trail",
    units: [
      {
        unit: "110",
        currentRent: 1100,
        monthlyRent: 1100,
        lotRent: 350,
        cost: 30000,
        terms: {
          "7": { totalPI: 63000, taxesInsurance: 12600 },
          "9": { totalPI: 81000, taxesInsurance: 16200 },
          "11": { totalPI: 99000, taxesInsurance: 19800 },
          "13": { totalPI: 117000, taxesInsurance: 23400 },
        },
      },
      {
        unit: "115",
        currentRent: 895,
        monthlyRent: 895,
        lotRent: 350,
        cost: 30000,
        terms: {
          "7": { totalPI: 45780, taxesInsurance: 12600 },
          "9": { totalPI: 58860, taxesInsurance: 16200 },
          "11": { totalPI: 71940, taxesInsurance: 19800 },
          "13": { totalPI: 85020, taxesInsurance: 23400 },
        },
      },
    ],
  },
  {
    park: "Hudson Haven Estates MH Park",
    units: [
      {
        unit: "102",
        currentRent: 750,
        monthlyRent: 750,
        lotRent: 400,
        cost: 15000,
        terms: {
          "7": { totalPI: 29400, taxesInsurance: 12600 },
          "9": { totalPI: 37800, taxesInsurance: 16200 },
          "11": { totalPI: 46200, taxesInsurance: 19800 },
          "13": { totalPI: 54600, taxesInsurance: 23400 },
        },
      },
      {
        unit: "103",
        currentRent: 925,
        monthlyRent: 925,
        lotRent: 400,
        cost: 25000,
        terms: {
          "7": { totalPI: 44100, taxesInsurance: 12600 },
          "9": { totalPI: 56700, taxesInsurance: 16200 },
          "11": { totalPI: 69300, taxesInsurance: 19800 },
          "13": { totalPI: 81900, taxesInsurance: 23400 },
        },
      },
    ],
  },
];

const PARK_DEFAULTS: Record<string, { lotRent: number; taxesInsuranceByTerm: Record<string, number> }> = {
  "Breeze Pointe MHP": { lotRent: 350, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
  "Fox Tail Trail": { lotRent: 350, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
  "Hudson Haven Estates MH Park": { lotRent: 400, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
  "Hudson MH Park": { lotRent: 400, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
  "Linda Lane MH Park": { lotRent: 400, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
  "Oak Grove Mobile Home Community": { lotRent: 400, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
  "Southview MH Park": { lotRent: 400, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
  "Virginia Avenue MH Park": { lotRent: 300, taxesInsuranceByTerm: { "7": 12600, "9": 16200, "11": 19800, "13": 23400 } },
};

function currency(value: number | string | null | undefined) {
  const number = Number(value || 0);
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(number);
}

function displayProfit(value: number | string | null | undefined) {
  if (value === "-" || value === null || value === undefined) return "-";
  return currency(value);
}

function sortUnits(items: Array<{ unit: string }>) {
  return [...items].sort((a, b) => {
    const aText = String(a.unit).trim();
    const bText = String(b.unit).trim();
    const aNum = /^\d+$/.test(aText) ? Number(aText) : null;
    const bNum = /^\d+$/.test(bText) ? Number(bText) : null;
    if (aNum !== null && bNum !== null) return aNum - bNum;
    if (aNum !== null) return -1;
    if (bNum !== null) return 1;
    return aText.localeCompare(bText, undefined, { numeric: true, sensitivity: "base" });
  });
}

function computeTermsFromRent(
  monthlyRent: number,
  lotRent: number,
  taxesInsuranceByTerm: Record<string, number>,
  maxPaymentDelta = 100,
) {
  const terms = ["7", "9", "11", "13"];
  const out: Record<string, { totalPI: number; taxesInsurance: number; profit: number }> = {};

  terms.forEach((t) => {
    const ti = Number(taxesInsuranceByTerm?.[t] || 0);
    const monthlyTI = ti / (Number(t) * 12);
    const monthlyHomePayment = Math.max(Number(monthlyRent || 0) - Number(lotRent || 0) - monthlyTI, 0);
    const rawMonthlyTotal = Number(lotRent || 0) + monthlyTI + monthlyHomePayment;
    const cappedMonthlyTotal = Math.min(rawMonthlyTotal, Number(monthlyRent || 0) + Number(maxPaymentDelta || 0));
    const cappedHomePayment = Math.max(cappedMonthlyTotal - Number(lotRent || 0) - monthlyTI, 0);
    out[t] = {
      totalPI: cappedHomePayment * Number(t) * 12,
      taxesInsurance: ti,
      profit: 0,
    };
  });

  return out;
}

function InputField({
  label,
  value,
  onChange,
  type = "text",
  readOnly = false,
  placeholder = "",
}: {
  label: string;
  value: string | number;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  type?: string;
  readOnly?: boolean;
  placeholder?: string;
}) {
  return (
    <label className="block text-sm text-slate-700">
      <div className="mb-1 font-medium">{label}</div>
      <input
        type={type}
        readOnly={readOnly}
        placeholder={placeholder}
        className={`h-11 w-full rounded-2xl border border-slate-300 bg-white px-3 text-sm outline-none transition focus:border-slate-500 ${readOnly ? "bg-slate-100 text-slate-500" : ""}`}
        value={value}
        onChange={onChange}
      />
    </label>
  );
}

function Panel({
  title,
  icon,
  children,
  right,
}: {
  title?: string;
  icon?: React.ReactNode;
  children: React.ReactNode;
  right?: React.ReactNode;
}) {
  return (
    <div className="rounded-3xl border border-slate-200 bg-white shadow-sm">
      {(title || right) && (
        <div className="flex items-center justify-between gap-3 border-b border-slate-100 px-6 py-4">
          <div className="flex items-center gap-2 text-xl font-semibold text-slate-900">
            {icon}
            {title}
          </div>
          {right}
        </div>
      )}
      <div className="p-6">{children}</div>
    </div>
  );
}

function MetricCard({ label, value, strong = false }: { label: string; value: string; strong?: boolean }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4">
      <div className="text-sm text-slate-500">{label}</div>
      <div className={`mt-2 break-words text-2xl font-semibold tracking-tight ${strong ? "text-slate-950" : "text-slate-800"}`}>
        {value}
      </div>
    </div>
  );
}

function parseWorkbook(
  file: File,
  maxPaymentDelta = 100,
): Promise<{
  parkData: ParkRow[];
  defaultsMap: Record<string, { lotRent: number; taxesInsuranceByTerm: Record<string, number> }>;
}> {
  const reader = new FileReader();

  return new Promise((resolve, reject) => {
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const homesSheet = wb.Sheets["Homes"];
        const defaultsSheet = wb.Sheets["Park_Defaults"];

        if (!homesSheet || !defaultsSheet) {
          reject(new Error("Workbook must include Homes and Park_Defaults sheets."));
          return;
        }

        const homesRows = XLSX.utils.sheet_to_json<Record<string, any>>(homesSheet, { defval: "" });
        const defaultRows = XLSX.utils.sheet_to_json<Record<string, any>>(defaultsSheet, { defval: "" });

        const defaultsMap: Record<string, { lotRent: number; taxesInsuranceByTerm: Record<string, number> }> = {};
        defaultRows.forEach((row) => {
          const park = String(row.Park || row.park || "").trim();
          if (!park) return;
          defaultsMap[park] = {
            lotRent: Number(row.Lot_Rent ?? row["Lot Rent"] ?? row.lotRent ?? 0),
            taxesInsuranceByTerm: {
              "7": Number(row.TI_7 ?? row["TI 7"] ?? 0),
              "9": Number(row.TI_9 ?? row["TI 9"] ?? 0),
              "11": Number(row.TI_11 ?? row["TI 11"] ?? 0),
              "13": Number(row.TI_13 ?? row["TI 13"] ?? 0),
            },
          };
        });

        const grouped: Record<string, UnitRow[]> = {};
        homesRows.forEach((row) => {
          const park = String(row.Park || row.park || "").trim();
          const unit = String(row.Unit || row.unit || "").trim();
          if (!park || !unit) return;

          const defaults = defaultsMap[park] || {
            lotRent: 0,
            taxesInsuranceByTerm: { "7": 0, "9": 0, "11": 0, "13": 0 },
          };

          const monthlyRent = Number(row.Monthly_Rent ?? row["Monthly Rent"] ?? row["Monthly Rent ($/mo)"] ?? row.monthlyRent ?? 0);
          const cost = Number(row.Cost ?? row["Cost ($)"] ?? row.cost ?? 0);

          if (!grouped[park]) grouped[park] = [];
          grouped[park].push({
            unit,
            currentRent: Number(row.Current_Rent ?? row["Current Rent"] ?? row.currentRent ?? monthlyRent),
            monthlyRent,
            lotRent: defaults.lotRent,
            cost,
            terms: computeTermsFromRent(monthlyRent, defaults.lotRent, defaults.taxesInsuranceByTerm, maxPaymentDelta),
          });
        });

        const parkData = Object.entries(grouped).map(([park, units]) => ({ park, units }));
        resolve({ parkData, defaultsMap });
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () => reject(new Error("Could not read file."));
    reader.readAsArrayBuffer(file);
  });
}

function parseGoogleVisualizationResponse(rawText: string) {
  const start = rawText.indexOf("(");
  const end = rawText.lastIndexOf(");");
  if (start < 0 || end < 0) {
    throw new Error("Google Sheets response was not in the expected format.");
  }
  const jsonText = rawText.slice(start + 1, end);
  const parsed = JSON.parse(jsonText);
  const cols = parsed.table.cols.map((col: any) => col.label || col.id || "");
  const rows = parsed.table.rows.map((row: any) => {
    const out: Record<string, any> = {};
    cols.forEach((col: string, idx: number) => {
      out[col] = row.c?.[idx]?.v ?? "";
    });
    return out;
  });
  return rows;
}

async function fetchGoogleSheetsDataset(
  homesUrl: string,
  defaultsUrl: string,
  maxPaymentDelta = 100,
): Promise<{
  parkData: ParkRow[];
  defaultsMap: Record<string, { lotRent: number; taxesInsuranceByTerm: Record<string, number> }>;
}> {
  const [homesRes, defaultsRes] = await Promise.all([fetch(homesUrl), fetch(defaultsUrl)]);
  if (!homesRes.ok || !defaultsRes.ok) {
    throw new Error("Could not load one or both Google Sheets tabs. Make sure both tabs are published to the web.");
  }

  const homesRows = parseGoogleVisualizationResponse(await homesRes.text());
  const defaultRows = parseGoogleVisualizationResponse(await defaultsRes.text());

  const defaultsMap: Record<string, { lotRent: number; taxesInsuranceByTerm: Record<string, number> }> = {};
  defaultRows.forEach((row) => {
    const park = String(row.Park || row.park || "").trim();
    if (!park) return;
    defaultsMap[park] = {
      lotRent: Number(row.Lot_Rent ?? row["Lot Rent"] ?? row.lotRent ?? 0),
      taxesInsuranceByTerm: {
        "7": Number(row.TI_7 ?? row["TI 7"] ?? 0),
        "9": Number(row.TI_9 ?? row["TI 9"] ?? 0),
        "11": Number(row.TI_11 ?? row["TI 11"] ?? 0),
        "13": Number(row.TI_13 ?? row["TI 13"] ?? 0),
      },
    };
  });

  const grouped: Record<string, UnitRow[]> = {};
  homesRows.forEach((row) => {
    const park = String(row.Park || row.park || "").trim();
    const unit = String(row.Unit || row.unit || "").trim();
    if (!park || !unit) return;

    const defaults = defaultsMap[park] || {
      lotRent: 0,
      taxesInsuranceByTerm: { "7": 0, "9": 0, "11": 0, "13": 0 },
    };

    const monthlyRent = Number(row.Monthly_Rent ?? row["Monthly Rent"] ?? row.monthlyRent ?? 0);
    const cost = Number(row.Cost ?? row.cost ?? 0);

    if (!grouped[park]) grouped[park] = [];
    grouped[park].push({
      unit,
      currentRent: Number(row.Current_Rent ?? row["Current Rent"] ?? row.currentRent ?? monthlyRent),
      monthlyRent,
      lotRent: defaults.lotRent,
      cost,
      terms: computeTermsFromRent(monthlyRent, defaults.lotRent, defaults.taxesInsuranceByTerm, maxPaymentDelta),
    });
  });

  const parkData = Object.entries(grouped).map(([park, units]) => ({ park, units }));
  return { parkData, defaultsMap };
}

export default function WebsiteUIMobileHomeFinancing() {
  const [parkData, setParkData] = useState<ParkRow[]>(REAL_PARK_DATA);
  const [selectedPark, setSelectedPark] = useState(REAL_PARK_DATA[0]?.park || "");
  const [selectedUnit, setSelectedUnit] = useState(REAL_PARK_DATA[0]?.units?.[0]?.unit || "");
  const [selectedTerm, setSelectedTerm] = useState("13");
  const [editMode, setEditMode] = useState(false);
  const [draftHome, setDraftHome] = useState<any>(null);
  const [maxPaymentDelta, setMaxPaymentDelta] = useState(100);
  const [uploadStatus, setUploadStatus] = useState(
    "Built-in dataset is only a starter sample. Upload your workbook to load the full property list.",
  );
  const [customerView, setCustomerView] = useState(false);
  const [customerName, setCustomerName] = useState("");
  const [downPayment, setDownPayment] = useState(0);
  const [downPaymentOverride, setDownPaymentOverride] = useState(false);
  const [preparedBy, setPreparedBy] = useState("Rene");
  const [dealDate, setDealDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [googleHomesUrl, setGoogleHomesUrl] = useState(() => localStorage.getItem(GOOGLE_HOMES_URL_KEY) || "");
  const [googleDefaultsUrl, setGoogleDefaultsUrl] = useState(() => localStorage.getItem(GOOGLE_DEFAULTS_URL_KEY) || "");
  const [isSyncingSheets, setIsSyncingSheets] = useState(false);

  useEffect(() => {
    document.title = "Valera Properties - Pricing Tool";
  }, []);

  const parkNames = useMemo(() => parkData.map((park) => park.park), [parkData]);

  const parkMap = useMemo(() => {
    const map: Record<string, UnitRow[]> = {};
    for (const park of parkData) map[park.park] = sortUnits(park.units);
    return map;
  }, [parkData]);

  const units = useMemo(() => parkMap[selectedPark] || [], [parkMap, selectedPark]);
  const totalUnits = useMemo(() => parkData.reduce((sum, park) => sum + park.units.length, 0), [parkData]);
  const current = useMemo(() => units.find((item) => item.unit === selectedUnit) || units[0] || null, [units, selectedUnit]);

  const availableTerms = useMemo(() => {
    if (!current?.terms) return [] as Array<{ term: string; data: { totalPI: number; taxesInsurance: number }; profit: number }>;
    return Object.entries(current.terms)
      .map(([term, data]) => ({ term, data, profit: Number(data?.totalPI || 0) - Number(current?.cost || 0) }))
      .filter((item) => item.profit >= 0)
      .sort((a, b) => Number(a.term) - Number(b.term));
  }, [current]);

  const bestTerm = useMemo(() => {
    if (!availableTerms.length) return "";
    return [...availableTerms].sort((a, b) => b.profit - a.profit)[0]?.term || availableTerms[0].term;
  }, [availableTerms]);

  const activeTerm = availableTerms.some((item) => item.term === selectedTerm) ? selectedTerm : bestTerm;
  const termData = useMemo(() => {
    if (!activeTerm) return { totalPI: 0, taxesInsurance: 0 };
    return current?.terms?.[activeTerm] || { totalPI: 0, taxesInsurance: 0 };
  }, [current, activeTerm]);

  const totalPIPaidByBuyer = Number(termData.totalPI || 0);
  const monthlyTaxesInsurance = activeTerm ? Number(termData.taxesInsurance || 0) / (Number(activeTerm) * 12) : 0;
  const monthlyHomePayment = Math.max(Number(current?.monthlyRent || 0) - Number(current?.lotRent || 0) - monthlyTaxesInsurance, 0);
  const currentRentDisplay = Number(current?.currentRent ?? current?.monthlyRent ?? 0);
  const interestPaidDuringTerm = Math.max(totalPIPaidByBuyer - Number(current?.cost || 0), 0);
  const profitValue = totalPIPaidByBuyer - Number(current?.cost || 0);
  const salespersonCommission = profitValue > 0 ? profitValue * 0.025 : 0;
  const totalNewMonthlyPayment = Number(current?.monthlyRent || 0);

  const customerDeltaLabel = useMemo(() => {
    const diff = (current?.monthlyRent || 0) - currentRentDisplay;
    if (diff === 0) return "Same monthly payment";
    if (diff > 0) return `Only ${currency(diff)} more per month`;
    return `Save ${currency(Math.abs(diff))} per month`;
  }, [current, currentRentDisplay]);

  const suggestedDownPayment = useMemo(() => {
    const cost = Number(current?.cost || 0);
    if (cost >= 40000) return 3500;
    if (cost >= 20000) return 2500;
    return 1500;
  }, [current]);

  const dealReference = useMemo(() => {
    const parkCode = (selectedPark || "PARK").replace(/[^A-Za-z0-9]/g, "").slice(0, 6).toUpperCase() || "PARK";
    const unitCode = String(current?.unit || "UNIT").replace(/[^A-Za-z0-9]/g, "").slice(0, 6).toUpperCase() || "UNIT";
    const dateCode = (dealDate || new Date().toISOString().slice(0, 10)).replace(/-/g, "");
    return `${parkCode}-${unitCode}-${dateCode}`;
  }, [selectedPark, current, dealDate]);

  const validThroughDate = useMemo(() => {
    const base = dealDate ? new Date(`${dealDate}T00:00:00`) : new Date();
    const next = new Date(base);
    next.setDate(next.getDate() + 15);
    return next.toISOString().slice(0, 10);
  }, [dealDate]);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length) {
          setParkData(parsed);
          setSelectedPark(parsed[0]?.park || "");
          setSelectedUnit(parsed[0]?.units?.[0]?.unit || "");
          setUploadStatus(`Auto-loaded ${parsed.reduce((s: number, p: ParkRow) => s + p.units.length, 0)} homes from last uploaded dataset.`);
        }
      }
    } catch (err) {
      console.error("Failed to auto-load saved dataset", err);
    }
  }, []);

  useEffect(() => {
    if (!units.some((item) => item.unit === selectedUnit) && units[0]?.unit) {
      setSelectedUnit(units[0].unit);
    }
  }, [units, selectedUnit]);

  useEffect(() => {
    if (!availableTerms.length) return;
    const isCurrentValid = availableTerms.some((item) => item.term === selectedTerm);
    if (!isCurrentValid) setSelectedTerm(bestTerm);
  }, [availableTerms, bestTerm, selectedTerm]);

  useEffect(() => {
    if (!downPaymentOverride) setDownPayment(suggestedDownPayment);
  }, [suggestedDownPayment, downPaymentOverride]);

  useEffect(() => {
    if (!editMode || !draftHome) return;
    const defaults = PARK_DEFAULTS[draftHome.park as keyof typeof PARK_DEFAULTS] || {
      lotRent: draftHome.lotRent || 0,
      taxesInsuranceByTerm: {
        "7": draftHome?.terms?.["7"]?.taxesInsurance || 0,
        "9": draftHome?.terms?.["9"]?.taxesInsurance || 0,
        "11": draftHome?.terms?.["11"]?.taxesInsurance || 0,
        "13": draftHome?.terms?.["13"]?.taxesInsurance || 0,
      },
    };
    setDraftHome((prev: any) => ({
      ...prev,
      terms: computeTermsFromRent(prev.monthlyRent, prev.lotRent, defaults.taxesInsuranceByTerm, maxPaymentDelta),
    }));
  }, [maxPaymentDelta, editMode]);

  const handleUploadWorkbook = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setUploadStatus("Loading workbook...");
      const { parkData: uploadedData } = await parseWorkbook(file, maxPaymentDelta);
      if (!uploadedData.length) throw new Error("No homes found in workbook.");
      setParkData(uploadedData);
      setSelectedPark(uploadedData[0]?.park || "");
      setSelectedUnit(uploadedData[0]?.units?.[0]?.unit || "");
      localStorage.setItem(STORAGE_KEY, JSON.stringify(uploadedData));
      setUploadStatus(`Loaded ${uploadedData.reduce((s, p) => s + p.units.length, 0)} homes from workbook.`);
      setEditMode(false);
      setDraftHome(null);
    } catch (err: any) {
      setUploadStatus(err?.message || "Upload failed.");
    } finally {
      event.target.value = "";
    }
  };

  const syncFromGoogleSheets = async () => {
    if (!googleHomesUrl || !googleDefaultsUrl) {
      setUploadStatus("Enter both Google Sheets published URLs first.");
      return;
    }
    try {
      setIsSyncingSheets(true);
      setUploadStatus("Syncing from Google Sheets...");
      localStorage.setItem(GOOGLE_HOMES_URL_KEY, googleHomesUrl);
      localStorage.setItem(GOOGLE_DEFAULTS_URL_KEY, googleDefaultsUrl);
      const { parkData: syncedData } = await fetchGoogleSheetsDataset(googleHomesUrl, googleDefaultsUrl, maxPaymentDelta);
      if (!syncedData.length) throw new Error("No homes were returned from Google Sheets.");
      setParkData(syncedData);
      setSelectedPark(syncedData[0]?.park || "");
      setSelectedUnit(syncedData[0]?.units?.[0]?.unit || "");
      localStorage.setItem(STORAGE_KEY, JSON.stringify(syncedData));
      setUploadStatus(`Synced ${syncedData.reduce((s, p) => s + p.units.length, 0)} homes from Google Sheets.`);
      setEditMode(false);
      setDraftHome(null);
    } catch (err: any) {
      setUploadStatus(err?.message || "Google Sheets sync failed.");
    } finally {
      setIsSyncingSheets(false);
    }
  };

  const updateDraftField = (field: string, value: string | number) => {
    setDraftHome((prev: any) => {
      const next = { ...prev, [field]: value };
      const defaults = PARK_DEFAULTS[next.park as keyof typeof PARK_DEFAULTS] || {
        lotRent: next.lotRent || 0,
        taxesInsuranceByTerm: {
          "7": next?.terms?.["7"]?.taxesInsurance || 0,
          "9": next?.terms?.["9"]?.taxesInsurance || 0,
          "11": next?.terms?.["11"]?.taxesInsurance || 0,
          "13": next?.terms?.["13"]?.taxesInsurance || 0,
        },
      };
      if (field === "park") next.lotRent = defaults.lotRent;
      if (field === "monthlyRent" || field === "park" || field === "lotRent") {
        next.terms = computeTermsFromRent(Number(next.monthlyRent || 0), Number(next.lotRent || 0), defaults.taxesInsuranceByTerm, maxPaymentDelta);
      }
      return next;
    });
  };

  const startAddNewHome = () => {
    const defaultPark = selectedPark || parkNames[0] || "";
    const defaults = PARK_DEFAULTS[defaultPark as keyof typeof PARK_DEFAULTS] || {
      lotRent: 0,
      taxesInsuranceByTerm: { "7": 0, "9": 0, "11": 0, "13": 0 },
    };
    const monthlyRent = 0;
    setDraftHome({
      park: defaultPark,
      originalUnit: "",
      unit: "",
      currentRent: monthlyRent,
      monthlyRent,
      lotRent: defaults.lotRent,
      cost: 0,
      terms: computeTermsFromRent(monthlyRent, defaults.lotRent, defaults.taxesInsuranceByTerm, maxPaymentDelta),
    });
    setEditMode(true);
  };

  const startEditCurrentHome = () => {
    if (!current) return;
    const defaults = PARK_DEFAULTS[selectedPark as keyof typeof PARK_DEFAULTS] || {
      lotRent: current.lotRent || 0,
      taxesInsuranceByTerm: {
        "7": current?.terms?.["7"]?.taxesInsurance || 0,
        "9": current?.terms?.["9"]?.taxesInsurance || 0,
        "11": current?.terms?.["11"]?.taxesInsurance || 0,
        "13": current?.terms?.["13"]?.taxesInsurance || 0,
      },
    };
    setDraftHome({
      park: selectedPark,
      originalUnit: current.unit,
      unit: current.unit,
      currentRent: Number(current.currentRent ?? current.monthlyRent ?? 0),
      monthlyRent: Number(current.monthlyRent || 0),
      lotRent: Number(current.lotRent || defaults.lotRent || 0),
      cost: Number(current.cost || 0),
      terms: computeTermsFromRent(Number(current.monthlyRent || 0), Number(current.lotRent || defaults.lotRent || 0), defaults.taxesInsuranceByTerm, maxPaymentDelta),
    });
    setEditMode(true);
  };

  const saveDraftHome = () => {
    if (!draftHome?.park || !draftHome?.unit) return;
    setParkData((prev) => {
      const next = prev.map((park) => ({ ...park, units: [...park.units] }));
      const parkIndex = next.findIndex((park) => park.park === draftHome.park);
      if (parkIndex < 0) return prev;
      const unitIndex = next[parkIndex].units.findIndex((u) => u.unit === draftHome.originalUnit);
      const cleanHome: UnitRow = {
        unit: draftHome.unit,
        currentRent: Number(draftHome.currentRent ?? draftHome.monthlyRent ?? 0),
        monthlyRent: Number(draftHome.monthlyRent || 0),
        lotRent: Number(draftHome.lotRent || 0),
        cost: Number(draftHome.cost || 0),
        terms: draftHome.terms,
      };
      if (unitIndex >= 0) next[parkIndex].units[unitIndex] = cleanHome;
      else next[parkIndex].units.push(cleanHome);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
      return next;
    });
    setSelectedPark(draftHome.park);
    setSelectedUnit(draftHome.unit);
    setEditMode(false);
    setDraftHome(null);
  };

  const removeCurrentHome = () => {
    if (!current || !selectedPark) return;
    setParkData((prev) => {
      const next = prev.map((park) => ({ ...park, units: [...park.units] }));
      const parkIndex = next.findIndex((park) => park.park === selectedPark);
      if (parkIndex < 0) return prev;
      next[parkIndex].units = next[parkIndex].units.filter((u) => u.unit !== current.unit);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
      return next;
    });
    setEditMode(false);
    setDraftHome(null);
  };

  const printDealSheet = () => {
    window.print();
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="mx-auto max-w-7xl px-4 py-8 md:px-8 lg:px-10">
        <div className="mb-8 flex flex-col gap-4 rounded-3xl border border-slate-200 bg-white p-6 shadow-sm md:flex-row md:items-start md:justify-between">
          <div className="flex items-center gap-4">
            <img src={LOGO_SRC} alt="Valera Properties" className="h-12 w-auto object-contain" />
            <div>
              <div className="mb-2 flex flex-wrap items-center gap-2">
                <span className="rounded-full bg-slate-900 px-3 py-1 text-xs font-semibold text-white">Website UI</span>
                <span className="rounded-full border border-slate-300 px-3 py-1 text-xs font-semibold text-slate-700">Deployment Ready</span>
              </div>
              <h1 className="text-2xl font-semibold tracking-tight">Mobile Home Payment Estimator</h1>
              <p className="mt-1 max-w-2xl text-sm text-slate-600">React + Vite friendly version for deployment to Vercel or Netlify.</p>
            </div>
          </div>

          <div className="flex flex-col gap-3 md:items-end">
            <div className="flex flex-wrap gap-2">
              <button type="button" className="inline-flex h-11 items-center rounded-2xl border border-slate-300 px-4 text-sm font-medium" onClick={startEditCurrentHome}><Pencil className="mr-2 h-4 w-4" />Edit Home</button>
              <button type="button" className="inline-flex h-11 items-center rounded-2xl border border-slate-300 px-4 text-sm font-medium" onClick={removeCurrentHome}><Trash2 className="mr-2 h-4 w-4" />Mark Sold / Remove</button>
              <button type="button" className="inline-flex h-11 items-center rounded-2xl bg-slate-900 px-4 text-sm font-medium text-white" onClick={startAddNewHome}><PlusCircle className="mr-2 h-4 w-4" />Add Home</button>
              <button type="button" className="inline-flex h-11 items-center rounded-2xl border border-slate-300 px-4 text-sm font-medium" onClick={printDealSheet}><Printer className="mr-2 h-4 w-4" />Print Deal Sheet</button>
              <button type="button" className="inline-flex h-11 items-center rounded-2xl border border-slate-300 px-4 text-sm font-medium" onClick={() => setCustomerView((prev) => !prev)}>
                {customerView ? <EyeOff className="mr-2 h-4 w-4" /> : <Eye className="mr-2 h-4 w-4" />}
                {customerView ? "Internal View" : "Customer View"}
              </button>
              <label className="inline-flex h-11 cursor-pointer items-center rounded-2xl border border-slate-300 px-4 text-sm font-medium">
                <Upload className="mr-2 h-4 w-4" />Upload Dataset
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUploadWorkbook} />
              </label>
              <button type="button" className="inline-flex h-11 items-center rounded-2xl border border-slate-300 px-4 text-sm font-medium disabled:opacity-60" onClick={syncFromGoogleSheets} disabled={isSyncingSheets}>
                <RefreshCw className={`mr-2 h-4 w-4 ${isSyncingSheets ? "animate-spin" : ""}`} />
                {isSyncingSheets ? "Syncing Sheets..." : "Sync Google Sheets"}
              </button>
            </div>
            {uploadStatus ? <div className="text-sm text-slate-600">{uploadStatus}</div> : null}
            <div className="grid gap-3 md:grid-cols-2 md:w-[640px]">
              <InputField label="Google Homes URL" value={googleHomesUrl} onChange={(e) => setGoogleHomesUrl(e.target.value)} placeholder="https://docs.google.com/...&sheet=Homes" />
              <InputField label="Google Defaults URL" value={googleDefaultsUrl} onChange={(e) => setGoogleDefaultsUrl(e.target.value)} placeholder="https://docs.google.com/...&sheet=Park_Defaults" />
            </div>
            <div className="text-xs text-slate-500 md:w-[640px]">Publish each Google Sheets tab to the web and paste the two gviz JSON URLs here: one for Homes and one for Park_Defaults.</div>
            <div className="grid grid-cols-2 gap-3 md:w-auto">
              <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3"><div className="text-xs text-slate-500">Parks Loaded</div><div className="mt-1 text-lg font-semibold">{parkNames.length}</div></div>
              <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3"><div className="text-xs text-slate-500">Total Units</div><div className="mt-1 text-lg font-semibold">{totalUnits}</div></div>
            </div>
          </div>
        </div>

        {editMode && draftHome && (
          <Panel title={draftHome.originalUnit ? "Edit Home" : "Add Home"} icon={<Save className="h-5 w-5" />}>
            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              <label className="block text-sm text-slate-700">
                <div className="mb-1 font-medium">Park</div>
                <select className="h-11 w-full rounded-2xl border border-slate-300 bg-white px-3 text-sm" value={draftHome.park} onChange={(e) => updateDraftField("park", e.target.value)}>
                  {parkNames.map((park) => <option key={park} value={park}>{park}</option>)}
                </select>
              </label>
              <InputField label="Unit" value={draftHome.unit} onChange={(e) => updateDraftField("unit", e.target.value)} />
              <InputField label="Current Rent" type="number" value={draftHome.currentRent ?? draftHome.monthlyRent} onChange={(e) => updateDraftField("currentRent", Number(e.target.value))} />
              <InputField label="Rent To Own Payment" type="number" value={draftHome.monthlyRent} onChange={(e) => updateDraftField("monthlyRent", Number(e.target.value))} />
              <InputField label="Lot Rent" type="number" readOnly value={draftHome.lotRent} onChange={() => {}} />
              <InputField label="Cost of Home" type="number" value={draftHome.cost} onChange={(e) => updateDraftField("cost", Number(e.target.value))} />
              <InputField label="Max Payment Delta" type="number" value={maxPaymentDelta} onChange={(e) => setMaxPaymentDelta(Number(e.target.value))} />
            </div>
            <div className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              {Object.keys(draftHome.terms || {}).map((term) => (
                <div key={term} className="rounded-2xl border border-slate-200 p-4">
                  <div className="mb-3 font-medium text-slate-900">{term} Year Term</div>
                  <InputField label="Total P&I" type="number" readOnly value={draftHome.terms[term]?.totalPI || 0} onChange={() => {}} />
                  <div className="mt-3"><InputField label="Taxes & Insurance" type="number" readOnly value={draftHome.terms[term]?.taxesInsurance || 0} onChange={() => {}} /></div>
                </div>
              ))}
            </div>
            <div className="mt-6 flex gap-3">
              <button type="button" className="inline-flex h-11 items-center rounded-2xl bg-slate-900 px-4 text-sm font-medium text-white" onClick={saveDraftHome}><Save className="mr-2 h-4 w-4" />Save Home</button>
              <button type="button" className="inline-flex h-11 items-center rounded-2xl border border-slate-300 px-4 text-sm font-medium" onClick={() => { setEditMode(false); setDraftHome(null); }}>Cancel</button>
            </div>
          </Panel>
        )}

        <div className={`grid gap-6 ${customerView ? "lg:grid-cols-1" : "lg:grid-cols-[360px_minmax(0,1fr)]"}`}>
          {!customerView && (
            <Panel title="Select Property" icon={<Home className="h-5 w-5" />}>
              <div className="space-y-5">
                <label className="block text-sm text-slate-700">
                  <div className="mb-1 font-medium">Park</div>
                  <select className="h-12 w-full rounded-2xl border border-slate-300 bg-white px-3 text-sm" value={selectedPark} onChange={(e) => setSelectedPark(e.target.value)}>
                    {parkNames.map((park) => <option key={park} value={park}>{park}</option>)}
                  </select>
                </label>
                <label className="block text-sm text-slate-700">
                  <div className="mb-1 font-medium">Home / Unit</div>
                  <select className="h-12 w-full rounded-2xl border border-slate-300 bg-white px-3 text-sm" value={selectedUnit} onChange={(e) => setSelectedUnit(e.target.value)}>
                    {units.map((item, idx) => <option key={`${item.unit}-${idx}`} value={item.unit}>{item.unit}</option>)}
                  </select>
                </label>
                <label className="block text-sm text-slate-700">
                  <div className="mb-1 font-medium">Term</div>
                  <select className="h-12 w-full rounded-2xl border border-slate-300 bg-white px-3 text-sm" value={activeTerm} onChange={(e) => setSelectedTerm(e.target.value)}>
                    {availableTerms.map((item) => <option key={item.term} value={item.term}>{item.term} Years</option>)}
                  </select>
                </label>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <div className="flex items-center gap-2 text-sm font-medium text-slate-700"><MapPin className="h-4 w-4" />Selected Park</div>
                  <div className="mt-2 text-lg font-semibold">{selectedPark || "—"}</div>
                  <div className="mt-1 text-sm text-slate-500">Unit {current?.unit || "—"} · {activeTerm || "—"} year term</div>
                </div>
              </div>
            </Panel>
          )}

          <div className="space-y-6">
            <Panel>
              <div className="flex items-center gap-2 text-sm font-medium text-slate-500">
                <DollarSign className="h-4 w-4" />
                {customerView ? "Rent To Own Payment" : "Monthly Rent"}
              </div>
              <div className="mt-3 text-5xl font-bold tracking-tight">{currency(current?.monthlyRent)}</div>
              {customerView ? (
                <>
                  <div className="mt-3 text-sm text-slate-600">
                    <div><span className="font-medium">Park:</span> {selectedPark}</div>
                    <div><span className="font-medium">Unit:</span> {current?.unit || "-"}</div>
                  </div>
                  <div className="mt-2 space-y-1 text-sm text-slate-600">
                    <div>Current Monthly Rent: <span className="font-medium">{currency(currentRentDisplay)}</span></div>
                    <div className="font-semibold text-emerald-600">{customerDeltaLabel}</div>
                  </div>
                  <p className="mt-3 text-sm text-slate-600">Simple customer-friendly pricing comparison.</p>
                  <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-4">
                    <InputField label="Customer Name" value={customerName} onChange={(e) => setCustomerName(e.target.value)} />
                    <label className="block text-sm text-slate-700">
                      <div className="mb-1 font-medium">Prepared By</div>
                      <select className="h-11 w-full rounded-2xl border border-slate-300 bg-white px-3 text-sm" value={preparedBy} onChange={(e) => setPreparedBy(e.target.value)}>
                        {["Rene", "Ino", "Aurora", "Christy"].map((name) => <option key={name} value={name}>{name}</option>)}
                      </select>
                    </label>
                    <InputField label="Date" type="date" value={dealDate} onChange={(e) => setDealDate(e.target.value)} />
                    <InputField label="Deal Reference" value={dealReference} readOnly onChange={() => {}} />
                    <div className="space-y-2">
                      <InputField label="Down Payment" type="number" value={downPayment} onChange={(e) => { setDownPayment(Number(e.target.value)); setDownPaymentOverride(true); }} />
                      <div className="text-xs text-slate-500">Auto-filled by tier: under $20,000 = $1,500 · $20,000–$39,999 = $2,500 · $40,000+ = $3,500</div>
                      <button type="button" className="text-xs font-medium text-slate-700 underline" onClick={() => { setDownPaymentOverride(false); setDownPayment(suggestedDownPayment); }}>Reset to suggested down payment</button>
                    </div>
                  </div>
                  <div className="mt-4 rounded-2xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">
                    <div className="font-semibold">This estimate is valid for 15 days only.</div>
                    <div className="mt-1">Valid through: <span className="font-medium">{validThroughDate}</span></div>
                  </div>
                </>
              ) : (
                <p className="mt-3 text-sm text-slate-600">Customer-facing monthly number based on the loaded park dataset.</p>
              )}
            </Panel>

            <Panel title={customerView ? "Pricing Summary" : "Payment Breakdown"} icon={customerView ? undefined : <Calculator className="h-5 w-5" />}>
              <div className={`grid gap-4 ${customerView ? "md:grid-cols-2" : "md:grid-cols-2 xl:grid-cols-3"}`}>
                {customerView ? (
                  <>
                    <MetricCard label="Rent To Own Payment" value={currency(current?.monthlyRent)} strong />
                    <MetricCard label="Park" value={selectedPark} />
                    <MetricCard label="Unit" value={current?.unit || "-"} />
                    <MetricCard label="Customer" value={customerName || "-"} />
                    <MetricCard label="Prepared By" value={preparedBy || "-"} />
                    <MetricCard label="Down Payment" value={currency(downPayment)} />
                    <MetricCard label="Total Price Including Down Payment" value={currency(Number(current?.cost || 0) + Number(downPayment || 0))} />
                    <MetricCard label="Date" value={dealDate || "-"} />
                    <MetricCard label="Valid Through" value={validThroughDate || "-"} />
                    <MetricCard label="Deal Reference" value={dealReference} />
                    <MetricCard label="Current Monthly Rent" value={currency(currentRentDisplay)} />
                    <MetricCard label="Monthly Home Payment" value={currency(monthlyHomePayment)} />
                    <MetricCard label="Monthly Taxes & Insurance" value={currency(monthlyTaxesInsurance)} />
                  </>
                ) : (
                  <>
                    <MetricCard label="Lot Rent" value={currency(current?.lotRent)} />
                    <MetricCard label="Taxes & Insurance" value={currency(termData?.taxesInsurance)} />
                    <MetricCard label="Monthly Taxes & Insurance" value={currency(monthlyTaxesInsurance)} />
                    <MetricCard label="Monthly Home Payment" value={currency(monthlyHomePayment)} />
                    <MetricCard label="Total New Monthly Payment" value={currency(totalNewMonthlyPayment)} strong />
                    <MetricCard label="Cost of Home" value={currency(current?.cost)} />
                    <MetricCard label="Total P&I Paid by Buyer" value={currency(totalPIPaidByBuyer)} />
                    <MetricCard label="Interest Paid During Term" value={currency(interestPaidDuringTerm)} />
                    <MetricCard label="Salesperson Commission (2.5%)" value={currency(salespersonCommission)} />
                    <MetricCard label="Selected Term" value={`${activeTerm || "-"} years`} />
                  </>
                )}
              </div>

              {!customerView && (
                <div className="mt-6 rounded-3xl border border-slate-200 p-5">
                  <div className="mb-2 flex items-center gap-2 text-sm font-medium text-slate-500">
                    {profitValue >= 0 ? <TrendingUp className="h-4 w-4" /> : <TrendingDown className="h-4 w-4" />}
                    Profit / Loss Summary
                  </div>
                  <div className={`text-3xl font-bold tracking-tight ${profitValue >= 0 ? "text-emerald-600" : "text-red-600"}`}>{displayProfit(profitValue)}</div>
                  <p className="mt-2 text-sm text-slate-600">Calculated as total P&amp;I paid by buyer minus cost of each home.</p>
                  <p className="mt-2 text-sm text-slate-600">Down payment is added on top of the home price for presentation only and does not change the payment calculations shown here.</p>
                  <p className="mt-2 text-sm text-slate-600">Use a separate compliant APR model for installment agreement disclosures.</p>
                  <p className="mt-2 text-sm text-slate-600">Max payment constraint is set to {currency(maxPaymentDelta)} above monthly rent when structuring new homes in the editor.</p>
                </div>
              )}
            </Panel>
          </div>
        </div>

        <div className="mt-10 flex items-center justify-center opacity-40">
          <img src={LOGO_SRC} alt="Valera Properties" className="h-8 w-auto" />
        </div>
      </div>
    </div>
  );
}

const __smokeTests = [
  computeTermsFromRent(1000, 400, { "7": 12600, "9": 16200, "11": 19800, "13": 23400 }, 100)["7"].totalPI >= 0,
  sortUnits([{ unit: "10" }, { unit: "2" }, { unit: "A" }])[0].unit === "2",
  computeTermsFromRent(900, 350, { "7": 12600, "9": 16200, "11": 19800, "13": 23400 }, 100)["13"].taxesInsurance === 23400,
  currency(1000) === "$1,000",
];
console.assert(__smokeTests.every(Boolean), "Smoke tests failed");
