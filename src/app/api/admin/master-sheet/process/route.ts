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

  const cleanupPaths = [
    paths.file1,
    paths.file2,
    paths.file3,
    paths.jsonFile,
    ...(paths.metadataCsvs ?? []),
  ];

  try {
    const [buf1, buf2, buf3, jsonBuf, ...metadataBufs] = await Promise.all([
      downloadMasterSheetInput(paths.file1),
      downloadMasterSheetInput(paths.file2),
      downloadMasterSheetInput(paths.file3),
      downloadMasterSheetInput(paths.jsonFile),
      ...((paths.metadataCsvs ?? []).map((p) => downloadMasterSheetInput(p))),
    ]);

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
    const message = error instanceof Error ? error.message : "Merge failed.";
    return NextResponse.json({ error: message }, { status: 422 });
  } finally {
    await deleteMasterSheetInputs(cleanupPaths);
  }
}
