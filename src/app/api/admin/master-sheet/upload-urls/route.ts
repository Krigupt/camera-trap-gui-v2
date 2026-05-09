import { randomUUID } from "crypto";
import { NextRequest, NextResponse } from "next/server";
import { auth } from "@clerk/nextjs/server";
import { isAdminUserId } from "@/lib/admin";
import {
  buildMasterSheetObjectPath,
  createMasterSheetUploadUrl,
} from "@/lib/gcp";

export const runtime = "nodejs";
export const maxDuration = 60;

type InputFile = {
  field: string;
  name: string;
  type?: string;
};

export async function POST(request: NextRequest) {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  if (!isAdminUserId(userId)) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  let body: { files?: InputFile[] };
  try {
    body = (await request.json()) as { files?: InputFile[] };
  } catch {
    return NextResponse.json({ error: "Invalid JSON body." }, { status: 400 });
  }

  const files = Array.isArray(body.files) ? body.files : [];
  if (!files.length) {
    return NextResponse.json({ error: "No files provided." }, { status: 400 });
  }

  const jobId = randomUUID();

  try {
    const urls = await Promise.all(
      files.map(async (file) => {
        const objectPath = buildMasterSheetObjectPath({
          jobId,
          field: file.field || "file",
          originalName: file.name || "file",
        });
        const signed = await createMasterSheetUploadUrl({
          objectPath,
          contentType: file.type || "application/octet-stream",
        });
        return {
          field: file.field,
          objectPath: signed.objectPath,
          bucketName: signed.bucketName,
          uploadUrl: signed.uploadUrl,
          contentType: file.type || "application/octet-stream",
        };
      })
    );

    return NextResponse.json({ jobId, files: urls });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Failed to create upload URLs.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
