"use client";

import { useState } from "react";
import { Download, Loader2 } from "lucide-react";

export function MasterSheetForm() {
  const [pending, setPending] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async (e: React.FormEvent<HTMLFormElement>) => {
  e.preventDefault();
  setError(null);

  const form = e.currentTarget;
  const fd = new FormData(form);

  setPending(true);

  try {
    const response = await fetch("/api/admin/master-sheet", {
      method: "POST",
      body: fd,
    });

    if (!response.ok) {
      let errorMessage = "Merge failed.";

      try {
        const err = await response.json();
        errorMessage = err.error || errorMessage;
      } catch {}

      throw new Error(errorMessage);
    }

    const blob = await response.blob();

    const url = window.URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "Master_AIxCT-3_Filled.csv";

    document.body.appendChild(a);
    a.click();
    a.remove();

    window.URL.revokeObjectURL(url);
  } catch (err) {
    setError(err instanceof Error ? err.message : "Merge failed.");
  } finally {
    setPending(false);
  }
};

  const fieldClass =
    "block w-full text-sm text-gray-900 file:mr-3 file:rounded-lg file:border-0 file:bg-indigo-50 file:px-3 file:py-2 file:text-sm file:font-medium file:text-indigo-800 hover:file:bg-indigo-100";

  return (
    <form
      onSubmit={submit}
      className="rounded-xl border border-gray-200 bg-white p-6 shadow-sm space-y-6"
    >
      <p className="text-sm text-gray-600">
        Upload the same inputs as your notebook merge: two Excel workbooks
        (Species sheet), species CSV, predictions JSON, and optional metadata
        CSVs for date/time lookup.
      </p>

      <div className="space-y-4">
        <label className="block">
          <span className="text-sm font-medium text-gray-800">
            File #1 — Human / AI / filenames (.xlsx)
          </span>
          <input
            name="file1"
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            required
            className={`mt-1 ${fieldClass}`}
          />
        </label>

        <label className="block">
          <span className="text-sm font-medium text-gray-800">
            File #2 — Incongruent reasons (.xlsx)
          </span>
          <input
            name="file2"
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            required
            className={`mt-1 ${fieldClass}`}
          />
        </label>

        <label className="block">
          <span className="text-sm font-medium text-gray-800">
            File #3 — filename + species (.csv)
          </span>
          <input
            name="file3"
            type="file"
            accept=".csv,text/csv"
            required
            className={`mt-1 ${fieldClass}`}
          />
        </label>

        <label className="block">
          <span className="text-sm font-medium text-gray-800">
            Predictions JSON (e.g. B1_result.json)
          </span>
          <input
            name="jsonFile"
            type="file"
            accept=".json,application/json"
            required
            className={`mt-1 ${fieldClass}`}
          />
        </label>

        <label className="block">
          <span className="text-sm font-medium text-gray-800">
            Optional — metadata CSVs (date/time per image)
          </span>
          <input
            name="metadataCsvs"
            type="file"
            accept=".csv,text/csv"
            multiple
            className={`mt-1 ${fieldClass}`}
          />
          <span className="mt-1 block text-xs text-gray-500">
            Same role as CSVs in your notebook&apos;s folder: columns with
            file/filename and timestamp/date/time.
          </span>
        </label>
      </div>

      {error ? (
        <p className="text-sm text-red-600 rounded-lg bg-red-50 px-3 py-2">
          {error}
        </p>
      ) : null}

      <button
        type="submit"
        disabled={pending}
        className="inline-flex items-center justify-center gap-2 rounded-lg bg-indigo-600 px-4 py-2.5 text-sm font-medium text-white hover:bg-indigo-700 disabled:opacity-50"
      >
        {pending ? (
          <Loader2 className="h-4 w-4 animate-spin" />
        ) : (
          <Download className="h-4 w-4" />
        )}
        Build &amp; download master CSV
      </button>
    </form>
  );
}