import { NextRequest, NextResponse } from "next/server";
import { auth } from "@clerk/nextjs/server";
import { isAdminUserId } from "@/lib/admin";
import { mergeMasterSheet } from "@/lib/master-sheet-merge";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

export async function POST(request: NextRequest) {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  if (!isAdminUserId(userId)) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  let formData: FormData;
  try {
    formData = await request.formData();
  } catch {
    return NextResponse.json({ error: "Invalid form data" }, { status: 400 });
  }

  const file1 = formData.get("file1");
  const file2 = formData.get("file2");
  const file3 = formData.get("file3");
  const jsonFile = formData.get("jsonFile");

  if (!(file1 instanceof File) || !file1.size) {
    return NextResponse.json(
      { error: "File #1 (.xlsx) is required." },
      { status: 400 }
    );
  }
  if (!(file2 instanceof File) || !file2.size) {
    return NextResponse.json(
      { error: "File #2 (.xlsx) is required." },
      { status: 400 }
    );
  }
  if (!(file3 instanceof File) || !file3.size) {
    return NextResponse.json(
      { error: "File #3 (.csv) is required." },
      { status: 400 }
    );
  }
  if (!(jsonFile instanceof File) || !jsonFile.size) {
    return NextResponse.json(
      { error: "Predictions JSON is required." },
      { status: 400 }
    );
  }

  const metadataBuffers = await Promise.all(
    formData
      .getAll("metadataCsvs")
      .filter(
        (entry): entry is File =>
          entry instanceof File && entry.size > 0
      )
      .map((file) => file.arrayBuffer())
  );

  try {
    const [buf1, buf2, buf3, jsonBuf] = await Promise.all([
      file1.arrayBuffer(),
      file2.arrayBuffer(),
      file3.arrayBuffer(),
      jsonFile.text(),
    ]);

    const csv = mergeMasterSheet({
      file1: buf1,
      file2: buf2,
      file3: buf3,
      jsonText: jsonBuf,
      metadataCsvBuffers:
        metadataBuffers.length > 0 ? metadataBuffers : undefined,
    });

    return new NextResponse(csv, {
      status: 200,
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition":
          'attachment; filename="Master_AIxCT-3_Filled.csv"',
      },
    });
  } catch (e) {
    const message = e instanceof Error ? e.message : "Merge failed.";
    return NextResponse.json({ error: message }, { status: 422 });
  }
}
