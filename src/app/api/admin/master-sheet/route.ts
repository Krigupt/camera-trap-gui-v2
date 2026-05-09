import { NextRequest, NextResponse } from "next/server";
import { auth } from "@clerk/nextjs/server";
import { isAdminUserId } from "@/lib/admin";
import { mergeMasterSheet } from "@/lib/master-sheet-merge";
import { Storage } from "@google-cloud/storage";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300; 

// ==========================================
// 1. Explicitly load Google Credentials from Vercel Env
// ==========================================
const storage = new Storage({
  projectId: process.env.GOOGLE_CLOUD_PROJECT_ID,
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    // The .replace() is absolutely critical for Vercel to parse the private key correctly
    private_key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
  },
});

const BUCKET_NAME = process.env.GCP_DEFAULT_BUCKET || "camera-trap-p-e4-2020";
const bucket = storage.bucket(BUCKET_NAME);

// Helper to convert Node Buffer to ArrayBuffer for your merge function
function toArrayBuffer(buf: Buffer): ArrayBuffer {
  const ab = new ArrayBuffer(buf.byteLength);
  new Uint8Array(ab).set(buf);
  return ab;
}

type ProcessBody = {
  objectPaths?: {
    file1?: string;
    file2?: string;
    file3?: string;
    jsonFile?: string;
    metadataCsvs?: string[];
  };
};

export async function POST(request: NextRequest) {
  let cleanupPaths: string[] = [];
  
  try {
    const { userId } = await auth();
    if (!userId) return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    if (!isAdminUserId(userId)) return NextResponse.json({ error: "Forbidden" }, { status: 403 });

    let body: ProcessBody;
    try {
      body = (await request.json()) as ProcessBody;
    } catch {
      return NextResponse.json({ error: "Invalid JSON body." }, { status: 400 });
    }

    const paths = body.objectPaths;
    if (!paths?.file1 || !paths.file2 || !paths.file3 || !paths.jsonFile) {
      return NextResponse.json({ error: "Missing required file paths." }, { status: 400 });
    }

    // Keep track of what we need to delete later
    cleanupPaths = [
      paths.file1,
      paths.file2,
      paths.file3,
      paths.jsonFile,
      ...((paths.metadataCsvs ?? []).filter(
        (p): p is string => typeof p === "string" && p.trim().length > 0
      )),
    ];

    // Helper to download directly from GCS to a Buffer
    const downloadFile = async (path: string): Promise<Buffer> => {
      const [buffer] = await bucket.file(path).download();
      return buffer;
    };

    // ==========================================
    // 2. Download all files from GCS into memory
    // ==========================================
    const [buf1, buf2, buf3, jsonBuf, ...metadataBufs] = await Promise.all([
      downloadFile(paths.file1),
      downloadFile(paths.file2),
      downloadFile(paths.file3),
      downloadFile(paths.jsonFile),
      ...((paths.metadataCsvs ?? []).map((p) => downloadFile(p))),
    ]);

    // ==========================================
    // 3. Process the merge
    // ==========================================
    const csv = mergeMasterSheet({
      file1: toArrayBuffer(buf1),
      file2: toArrayBuffer(buf2),
      file3: toArrayBuffer(buf3),
      jsonText: jsonBuf.toString("utf-8"),
      metadataCsvBuffers: metadataBufs.map((b) => toArrayBuffer(b)),
    });

    return new NextResponse(csv, {
      status: 200,
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition": 'attachment; filename="Master_AIxCT-3_Filled.csv"',
      },
    });

  } catch (error) {
    console.error("Master sheet process failed:", error);
    const message = error instanceof Error ? error.message : "Merge failed.";
    return NextResponse.json({ error: message }, { status: 500 });

  } finally {
    // ==========================================
    // 4. Cleanup temporary files from GCS bucket
    // ==========================================
    if (cleanupPaths.length > 0) {
      try {
        await Promise.all(
          cleanupPaths.map((path) => bucket.file(path).delete().catch(() => {}))
        );
      } catch (cleanupError) {
        console.error("Master sheet cleanup failed:", cleanupError);
      }
    }
  }
}