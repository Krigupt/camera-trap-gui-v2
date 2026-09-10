import { NextRequest, NextResponse } from 'next/server';
import connectDB from '@/lib/mongodb';
import ExcelData from '@/models/ExcelData';
import * as XLSX from 'xlsx';

export async function POST(request: NextRequest) {
  try {
    await connectDB();

    const { filename, bucketName } = await request.json();

    if (!filename) {
      return NextResponse.json({ error: 'Filename is required' }, { status: 400 });
    }
    if (!bucketName) {
      return NextResponse.json({ error: 'Bucket name is required' }, { status: 400 });
    }

    // Different batches (different years/sites) were often uploaded with the
    // exact same default spreadsheet filename (e.g. "detailed_species_P_E4_report.xlsx"
    // reused across 2017/2018/2020). Filtering by filename alone merges unrelated
    // batches together — bucketName reliably identifies the actual batch/year, so
    // scope by both.
    const allSheets = await ExcelData.find({ filename, bucketName });

    if (allSheets.length === 0) {
      return NextResponse.json({ error: 'No data found for this file' }, { status: 404 });
    }

    // Re-uploads of the same filename create multiple ExcelData documents that
    // share a sheetName. Tags are written to all of them (see /api/global-tags),
    // but a tagged image's row may only exist in one of the duplicates. Group by
    // sheetName and merge every duplicate's rows so no tagged image is dropped.
    const docsBySheetName = new Map<string, typeof allSheets>();
    for (const sheet of allSheets) {
      const arr = docsBySheetName.get(sheet.sheetName) || [];
      arr.push(sheet);
      docsBySheetName.set(sheet.sheetName, arr);
    }
    const sheetNames = [...docsBySheetName.keys()].sort((a, b) => a.localeCompare(b));

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Define tag columns
    const tagColumns = [
      'Blurry',
      'Low-light',
      'Body part',
      'Blends in',
      'Unidentifiable to taxonomic level by human ground-truth',
      'Other',
      'Similar species that does not occur in the area'
    ];

    // Process each sheet name, merging rows from every duplicate document
    for (const sheetName of sheetNames) {
      const docs = docsBySheetName.get(sheetName)!;

      // Group data by human-ai pairs
      const groupedData = new Map();

      for (const sheet of docs) {
        const sheetTagsByImage = sheet.sheetSpecificImageTags?.[sheetName] || {};

        sheet.data.forEach((row: any) => {
          const key = `${row.human}_vs_${row.ai}`;
          if (!groupedData.has(key)) {
            groupedData.set(key, {
              human: row.human,
              ai: row.ai,
              taggedImages: {
                'Blurry': new Set<string>(),
                'Low-light': new Set<string>(),
                'Body part': new Set<string>(),
                'Blends in': new Set<string>(),
                'Unidentifiable to taxonomic level by human ground-truth': new Set<string>(),
                'Other': new Set<string>(),
                'Similar species that does not occur in the area': new Set<string>()
              }
            });
          }

          const group = groupedData.get(key);

          // Add filenames to appropriate tag columns based on sheet-specific image tags
          if (row.imagePaths) {
            row.imagePaths.forEach((imagePath: string) => {
              // Escape dots in imagePath to match MongoDB storage format
              const escapedImagePath = imagePath.replace(/\./g, '\uff0e');
              // Get sheet-specific tags for this image (from its own document)
              const sheetTags = sheetTagsByImage[escapedImagePath];
              if (sheetTags && sheetTags.length > 0) {
                sheetTags.forEach((tag: string) => {
                  if (group.taggedImages[tag]) {
                    // Add this specific image to the tag column (use original imagePath for display)
                    group.taggedImages[tag].add(imagePath);
                  }
                });
              }
            });
          }
        });
      }

      // Convert to Excel format
      const worksheetData = [];
      
      // Add headers
      const headers = ['Human', 'AI', ...tagColumns, 'Notable images'];
      worksheetData.push(headers);
      
      // Add data rows
      groupedData.forEach((group) => {
        const row = [
          group.human,
          group.ai,
          ...tagColumns.map(tag => Array.from(group.taggedImages[tag]).join(', ')),
          '' // Notable images column (empty for now)
        ];
        worksheetData.push(row);
      });

      // Create worksheet
      const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
      
      // Set column widths
      const columnWidths = [
        { wch: 20 }, // Human
        { wch: 20 }, // AI
        { wch: 35 }, // Blurry
        { wch: 35 }, // Low-light
        { wch: 35 }, // Body part
        { wch: 35 }, // Blends in
        { wch: 65 }, // Unidentifiable...
        { wch: 35 }, // Other
        { wch: 45 }, // Similar species...
        { wch: 35 }  // Notable images
      ];
      worksheet['!cols'] = columnWidths;

      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    // Generate Excel buffer
    const excelBuffer = XLSX.write(workbook, { 
      type: 'buffer', 
      bookType: 'xlsx' 
    });

    // Create filename for download
    const downloadFilename = `${filename.replace('.xlsx', '')}_tagged_analysis.xlsx`;

    return new NextResponse(excelBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${downloadFilename}"`,
        'Cache-Control': 'no-cache',
      },
    });

  } catch (error) {
    console.error('Error exporting tagged data:', error);
    return NextResponse.json(
      { error: 'Failed to export tagged data' },
      { status: 500 }
    );
  }
}
