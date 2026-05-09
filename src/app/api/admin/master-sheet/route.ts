import { NextRequest, NextResponse } from "next/server";
import { auth } from "@clerk/nextjs/server";
import { isAdminUserId } from "@/lib/admin";
import { mergeMasterSheet } from "@/lib/master-sheet-merge";
import { Storage } from "@google-cloud/storage";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300; // Gives Vercel up to 5 mins for large merges

// Initialize GCS client. It automatically uses your environment credentials.
const storage = new Storage();

// Replace with your env variable if needed, using your known bucket as a fallback
const BUCKET_NAME = process.env.GCS_BUCKET_NAME || "camera-trap-p-e4-2020";
const bucket = storage.bucket(BUCKET_NAME);

export async function POST(request: NextRequest) {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  if (!isAdminUserId(userId)) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  let body;
  try {
    // FIX FOR THE 500 ERROR: Parsing JSON instead of FormData
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body" }, { status: 400 });
  }

  const { objectPaths } = body;
  if (!objectPaths) {
    return NextResponse.json({ error: "Missing objectPaths mapping" }, { status: 400 });
  }

  const { file1, file2, file3, jsonFile, metadataCsvs } = objectPaths;

  if (!file1 || !file2 || !file3 || !jsonFile) {
    return NextResponse.json(
      { error: "Missing required file paths." },
      { status: 400 }
    );
  }

  try {
    // Helper to download a GCS file and safely convert Node Buffer to standard ArrayBuffer
    const downloadToArrayBuffer = async (path: string): Promise<ArrayBuffer> => {
      const [buffer] = await bucket.file(path).download();
      return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
    };

    // Helper to download a GCS file as text (for JSON)
    const downloadToText = async (path: string): Promise<string> => {
      const [buffer] = await bucket.file(path).download();
      return buffer.toString("utf-8");
    };

    // Download all files concurrently from your bucket into backend memory
    const [buf1, buf2, buf3, jsonText] = await Promise.all([
      downloadToArrayBuffer(file1),
      downloadToArrayBuffer(file2),
      downloadToArrayBuffer(file3),
      downloadToText(jsonFile),
    ]);

    // Download any optional metadata files
    let metaBufs: ArrayBuffer[] = [];
    if (Array.isArray(metadataCsvs) && metadataCsvs.length > 0) {
      metaBufs = await Promise.all(metadataCsvs.map((path: string) => downloadToArrayBuffer(path)));
    }

    // Process the downloaded buffers using your existing logic
    const csv = mergeMasterSheet({
      file1: buf1,
      file2: buf2,
      file3: buf3,
      jsonText: jsonText,
      metadataCsvBuffers: metaBufs.length > 0 ? metaBufs : undefined,
    });

    // Send the merged CSV back to the frontend
    return new NextResponse(csv, {
      status: 200,
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition": 'attachment; filename="Master_AIxCT-3_Filled.csv"',
      },
    });
  } catch (e) {
    console.error("GCS Download or Merge error:", e);
    const message = e instanceof Error ? e.message : "Merge failed.";
    return NextResponse.json({ error: message }, { status: 422 });
  }
}