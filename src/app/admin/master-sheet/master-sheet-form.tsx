"use client";

import { useState } from "react";
import { Download, Loader2 } from "lucide-react";

export function MasterSheetForm() {
  const [pending, setPending] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const getFilenameFromContentDisposition = (value: string | null): string => {
    if (!value) return "Master_AIxCT-3_Filled.csv";
    const utf8Match = value.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8Match?.[1]) {
      try { return decodeURIComponent(utf8Match[1]); } catch { return utf8Match[1]; }
    }
    const simpleMatch = value.match(/filename="?([^"]+)"?/i);
    if (simpleMatch?.[1]) return simpleMatch[1];
    return "Master_AIxCT-3_Filled.csv";
  };

  const submit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setError(null);
    const form = e.currentTarget;
    const fd = new FormData(form);

    setPending(true);
    try {
      const file1 = fd.get("file1");
      const file2 = fd.get("file2");
      const file3 = fd.get("file3");
      const jsonFile = fd.get("jsonFile");

      if (!(file1 instanceof File) || !file1.size) throw new Error("File #1 (.xlsx) is required.");
      if (!(file2 instanceof File) || !file2.size) throw new Error("File #2 (.xlsx) is required.");
      if (!(file3 instanceof File) || !file3.size) throw new Error("File #3 (.csv) is required.");
      if (!(jsonFile instanceof File) || !jsonFile.size) throw new Error("Predictions JSON is required.");

      const metaEntries = fd.getAll("metadataCsvs");
      const metadataFiles = metaEntries.filter((item): item is File => item instanceof File && item.size > 0);

      const uploadSpec = [
        { field: "file1", file: file1 },
        { field: "file2", file: file2 },
        { field: "file3", file: file3 },
        { field: "jsonFile", file: jsonFile },
        ...metadataFiles.map((file, idx) => ({
          field: `metadataCsvs[${idx}]`,
          file,
        })),
      ];

      // 1. Get Signed URLs from your backend
      const signedResponse = await fetch("/api/admin/master-sheet/upload-urls", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          files: uploadSpec.map((item) => ({
            field: item.field,
            name: item.file.name,
            type: item.file.type || "application/octet-stream",
          })),
        }),
      });

      if (!signedResponse.ok) {
        const text = await signedResponse.text();
        throw new Error(`Failed to prepare uploads (HTTP ${signedResponse.status}). ${text.slice(0, 200)}`);
      }

      const signedBody = await signedResponse.json();
      const signedFiles = signedBody.files ?? [];
      const objectPathByField = new Map<string, string>();

      // 2. Upload files directly to Google Cloud Storage (Bypassing Vercel's 4.5MB limit)
      for (const item of uploadSpec) {
        const target = signedFiles.find((f: any) => f.field === item.field);
        if (!target) throw new Error(`Missing upload URL for ${item.field}.`);
        
        const put = await fetch(target.uploadUrl, {
          method: "PUT",
          headers: { "Content-Type": target.contentType },
          body: item.file,
        });

        if (!put.ok) {
          throw new Error(`Storage upload failed for ${item.file.name}. Check your bucket CORS policy!`);
        }
        objectPathByField.set(item.field, target.objectPath);
      }

      // 3. Tell the backend to process the uploaded files
      // NOTE: Make sure this path matches where you put the new `route.ts` we created!
      const processResponse = await fetch("/api/admin/master-sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          objectPaths: {
            file1: objectPathByField.get("file1"),
            file2: objectPathByField.get("file2"),
            file3: objectPathByField.get("file3"),
            jsonFile: objectPathByField.get("jsonFile"),
            metadataCsvs: metadataFiles
              .map((_, idx) => objectPathByField.get(`metadataCsvs[${idx}]`))
              .filter(Boolean),
          },
        }),
      });

      if (!processResponse.ok) {
        let errStr = `Merge failed (HTTP ${processResponse.status}).`;
        try {
          const errJson = await processResponse.json();
          errStr = errJson.error || errStr;
        } catch {
          errStr += await processResponse.text();
        }
        throw new Error(errStr);
      }

      // 4. Download the resulting CSV
      const blob = await processResponse.blob();
      const filename = getFilenameFromContentDisposition(processResponse.headers.get("Content-Disposition"));
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

    } catch (err) {
      setError(err instanceof Error ? err.message : "Merge failed.");
    } finally {
      setPending(false);
    }
  };

  const fieldClass = "block w-full text-sm text-gray-900 file:mr-3 file:rounded-lg file:border-0 file:bg-indigo-50 file:px-3 file:py-2 file:text-sm file:font-medium file:text-indigo-800 hover:file:bg-indigo-100";

  return (
    <form onSubmit={submit} className="rounded-xl border border-gray-200 bg-white p-6 shadow-sm space-y-6">
      <p className="text-sm text-gray-600">
        Upload the same inputs as your notebook merge: two Excel workbooks, species CSV, predictions JSON, and optional metadata CSVs.
      </p>

      <div className="space-y-4">
        <label className="block">
          <span className="text-sm font-medium text-gray-800">File #1 — Human / AI / filenames (.xlsx)</span>
          <input name="file1" type="file" accept=".xlsx" required className={`mt-1 ${fieldClass}`} />
        </label>
        <label className="block">
          <span className="text-sm font-medium text-gray-800">File #2 — Incongruent reasons (.xlsx)</span>
          <input name="file2" type="file" accept=".xlsx" required className={`mt-1 ${fieldClass}`} />
        </label>
        <label className="block">
          <span className="text-sm font-medium text-gray-800">File #3 — filename + species (.csv)</span>
          <input name="file3" type="file" accept=".csv" required className={`mt-1 ${fieldClass}`} />
        </label>
        <label className="block">
          <span className="text-sm font-medium text-gray-800">Predictions JSON</span>
          <input name="jsonFile" type="file" accept=".json" required className={`mt-1 ${fieldClass}`} />
        </label>
        <label className="block">
          <span className="text-sm font-medium text-gray-800">Optional — metadata CSVs</span>
          <input name="metadataCsvs" type="file" accept=".csv" multiple className={`mt-1 ${fieldClass}`} />
        </label>
      </div>

      {error && <p className="text-sm text-red-600 rounded-lg bg-red-50 px-3 py-2">{error}</p>}

      <button type="submit" disabled={pending} className="inline-flex items-center justify-center gap-2 rounded-lg bg-indigo-600 px-4 py-2.5 text-sm font-medium text-white hover:bg-indigo-700 disabled:opacity-50">
        {pending ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
        Build & download master CSV
      </button>
    </form>
  );
}