import { NextRequest, NextResponse } from "next/server";
import { auth } from "@clerk/nextjs/server";
import { isAdminUserId } from "@/lib/admin";
import { mergeMasterSheet } from "@/lib/master-sheet-merge";
import { Storage } from "@google-cloud/storage";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300; 

// ==========================================
// 1. Explicitly load Google Credentials safely
// ==========================================
let rawPrivateKey = process.env.GOOGLE_PRIVATE_KEY || "";
rawPrivateKey = rawPrivateKey.replace(/^"|"$/g, "");
rawPrivateKey = rawPrivateKey.replace(/\\n/g, "\n");

const storage = new Storage({
  projectId: process.env.GOOGLE_CLOUD_PROJECT_ID,
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: rawPrivateKey,
  },
});

const BUCKET_NAME = process.env.GCP_DEFAULT_BUCKET || "camera-trap-p-e4-2020";
const bucket = storage.bucket(BUCKET_NAME);

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

    cleanupPaths = [
      paths.file1,
      paths.file2,
      paths.file3,
      paths.jsonFile,
      ...((paths.metadataCsvs ?? []).filter(
        (p): p is string => typeof p === "string" && p.trim().length > 0
      )),
    ];

    // Helper: Returns a Buffer directly from GCS. No memory copying.
    const downloadFile = async (path: string): Promise<Buffer> => {
      const [buffer] = await bucket.file(path).download();
      return buffer;
    };

    // ==========================================
    // 2. Process the merge SEQUENTIALLY to save RAM
    // ==========================================
    // Instead of downloading everything at once, we pass functions so the 
    // merge utility can download, process, and delete files one at a time.
    const csv = await mergeMasterSheet({
      getFile1: () => downloadFile(paths.file1!),
      getFile2: () => downloadFile(paths.file2!),
      getFile3: () => downloadFile(paths.file3!),
      getJsonFile: () => downloadFile(paths.jsonFile!),
      getMetadataFiles: (paths.metadataCsvs || []).map((p) => () => downloadFile(p)),
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
    // 3. Cleanup temporary files from GCS bucket
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