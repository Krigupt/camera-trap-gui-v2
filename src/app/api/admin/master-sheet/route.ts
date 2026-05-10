import { NextRequest, NextResponse } from "next/server";
import { auth } from "@clerk/nextjs/server";
import { isAdminUserId } from "@/lib/admin";
import { mergeMasterSheet } from "@/lib/master-sheet-merge";
import {
  deleteMasterSheetInputs,
  downloadMasterSheetInput,
} from "@/lib/gcp";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

/**
 * Merge master-sheet inputs that were uploaded via signed URLs.
 *
 * IMPORTANT: Uses the same bucket as `/api/admin/master-sheet/upload-urls`
 * (`MASTER_SHEET_UPLOAD_BUCKET` or `GCP_DEFAULT_BUCKET` from `@/lib/gcp`).
 * Do not hardcode a different bucket here — that causes “bucket does not exist”
 * when uploads land in one bucket but this handler reads another.
 */

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
    if (!paths?.file1 || !paths?.file2 || !paths?.file3 || !paths?.jsonFile) {
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
        "Content-Disposition":
          'attachment; filename="Master_AIxCT-3_Filled.csv"',
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
        console.error("Master sheet cleanup failed:", cleanupError);
      }
    }
  }
}
