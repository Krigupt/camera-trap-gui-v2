import { NextRequest, NextResponse } from "next/server";
import { auth } from "@clerk/nextjs/server";
import { isAdminUserId } from "@/lib/admin";
import { mergeMasterSheet } from "@/lib/master-sheet-merge";
import {
  deleteMasterSheetInputs,
  downloadMasterSheetInput,
} from "@/lib/gcp";

export const runtime = "nodejs";
export const maxDuration = 300;

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
    if (!userId) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }
    if (!isAdminUserId(userId)) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }

    let body: ProcessBody;
    try {
      body = (await request.json()) as ProcessBody;
    } catch {
      return NextResponse.json({ error: "Invalid JSON body." }, { status: 400 });
    }

    const paths = body.objectPaths;
    if (!paths?.file1 || !paths.file2 || !paths.file3 || !paths.jsonFile) {
      return NextResponse.json(
        {
          error:
            "objectPaths.file1, file2, file3, and jsonFile are all required.",
        },
        { status: 400 }
      );
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

    // ==========================================
    // THE FIX: Sequential Merge Call
    // Instead of downloading everything with Promise.all, we pass the getter 
    // functions down so the merge script can download them one at a time.
    // Also added 'await' since the function is now asynchronous.
    // ==========================================
    const csv = await mergeMasterSheet({
      getFile1: () => downloadMasterSheetInput(paths.file1!),
      getFile2: () => downloadMasterSheetInput(paths.file2!),
      getFile3: () => downloadMasterSheetInput(paths.file3!),
      getJsonFile: () => downloadMasterSheetInput(paths.jsonFile!),
      getMetadataFiles: (paths.metadataCsvs ?? []).map(
        (p) => () => downloadMasterSheetInput(p)
      ),
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
    if (cleanupPaths.length > 0) {
      try {
        await deleteMasterSheetInputs(cleanupPaths);
      } catch (cleanupError) {
        // Cleanup must never mask the real merge error.
        console.error("Master sheet cleanup failed:", cleanupError);
      }
    }
  }
}