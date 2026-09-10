'use client';

import { Suspense, useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { useParams, useRouter, useSearchParams } from 'next/navigation';
import { ImageIcon, TagIcon, DownloadIcon, ArrowLeft } from 'lucide-react';

interface ExcelData {
  _id: string;
  filename: string;
  sheetName: string;
  bucketName: string;
  /** Present for uploads after uploadGroupId rollout — scopes CSV rows */
  uploadGroupId?: string;
  data: Array<{
    human: string;
    ai: string;
    filenames?: string;
    imagePaths?: string[];
    tags?: string[];
    isSelected?: boolean;
    imageTags?: { [imagePath: string]: string[] }; // Individual image tags
  }>;
  globalImageTags?: { [imagePath: string]: string[] }; // Global tags across all sheets (DEPRECATED)
  sheetSpecificImageTags?: { [sheetName: string]: { [imagePath: string]: string[] } }; // Sheet-specific tags
  globalImageSpecies?: { [imagePath: string]: string }; // Global species classifications across all sheets
}

interface Sheet {
  id: string;
  sheetName: string;
}

const TAGS = [
  'Blurry',
  'Low-light',
  'Body part',
  'Blends in',
  'Unidentifiable to taxonomic level by human ground-truth',
  'Other',
  'Similar species that does not occur in the area'
];

const SPECIES_LIST = [
  'Aves',
  'Mammalia',
  'NAN',
  'Black-tailed jackrabbit - Lepus californicus',
  'Bobcat - Lynx rufus',
  '(Desert cottontail) - Sylvilagus audubonii',
  'Coyote - Canis latrans',
  'Domestic horse - Equus ferus caballus',
  'Domestic cattle - Bos taurus',
  'Gray fox - Urocyon cinereoargenteus',
  'Mule deer - Odocoileus hemionus',
  'Northern raccoon - Procyon lotor',
  'Puma - Puma concolor',
  'Striped skunk - Mephitis mephitis',
  'Virginia opossum - Didelphis virginiana',
  'Black bear - Ursus americanus',
  'Wild boar - Sus scrofa',
  'Western spotted skunk - Spilogale gracilis',
  'Western gray squirrel - Sciurus griseus',
  'Eastern gray squirrel - Sciurus carolinensis',
  '(Dusky-footed woodrat) - Neotoma fuscipes',
  'Unknown - (no scientific name)'
];

// Helper function to get image paths from row data
const getImagePaths = (row: ExcelData['data'][0] | undefined | null): string[] => {
  if (!row) return []; // Add null/undefined check
  
  if (row.imagePaths && row.imagePaths.length > 0) {
    return row.imagePaths;
  }
  if (row.filenames) {
    return row.filenames.split(',').map(f => f.trim()).filter(Boolean);
  }
  return [];
};

// Fetch with timeout so serverless timeouts (e.g. Vercel "---") surface as errors
const fetchWithTimeout = (url: string, options: RequestInit = {}, timeoutMs = 25000): Promise<Response> => {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(url, {
    ...options,
    credentials: 'include',
    signal: controller.signal,
  }).finally(() => clearTimeout(timeout));
};

// Navbar component to avoid duplication
const Navbar = ({ selectedSheet, sheets, onSheetChange, onExport, isExporting, onBack, backLabel }: {
  selectedSheet: string;
  sheets: Sheet[];
  onSheetChange: (sheetId: string) => void;
  onExport: () => void;
  isExporting: boolean;
  onBack: () => void;
  backLabel: string;
}) => (
  <nav className="bg-white shadow-sm border-b">
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div className="flex justify-between h-16">
        <div className="flex items-center space-x-4">
          <button
            onClick={onBack}
            className="flex items-center space-x-2 text-gray-600 hover:text-gray-900 transition-colors"
          >
            <ArrowLeft className="w-5 h-5" />
            <span>{backLabel}</span>
          </button>
          <h1 className="text-xl font-semibold text-gray-900">Excel Data Dashboard</h1>
        </div>
        
        <div className="flex items-center space-x-4">
          <button
            onClick={onExport}
            disabled={isExporting}
            className="flex items-center space-x-2 bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition-colors"
          >
            {isExporting ? (
              <>
                <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                <span>Exporting...</span>
              </>
            ) : (
              <>
                <DownloadIcon className="w-4 h-4" />
                <span>Export Tagged Data</span>
              </>
            )}
          </button>
          
          <select 
            value={sheets.find(sheet => sheet.sheetName === selectedSheet)?.id || ''} 
            onChange={(e) => onSheetChange(e.target.value)}
            className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors border-none outline-none cursor-pointer"
          >
            <option value="" disabled>Select Taxonomic Level</option>
            {sheets.map((sheet) => (
              <option key={sheet.id} value={sheet.id} className="bg-white text-black">
                {sheet.sheetName}
              </option>
            ))}
          </select>
        </div>
      </div>
    </div>
  </nav>
);

// Image display component
const ImageDisplay = ({ 
  imagePaths, 
  currentIndex, 
  onIndexChange, 
  rowData,
  bucketName
}: {
  imagePaths: string[];
  currentIndex: number;
  onIndexChange: (index: number) => void;
  rowData: { human: string; ai: string };
  bucketName: string;
}) => {
  const [imageLoading, setImageLoading] = useState(false);
  const [imageUrl, setImageUrl] = useState<string>('#');

  const currentImagePath = imagePaths[currentIndex] || '';
  
  // Fetch the signed URL when the image path or bucket changes
  useEffect(() => {
    if (bucketName && currentImagePath) {
      setImageLoading(true);
      setImageUrl('#'); // Reset URL first
      
      fetch(`/api/images/${currentImagePath}?bucket=${bucketName}`)
        .then(response => response.json())
        .then(data => {
          if (data.success && data.imageUrl) {
            setImageUrl(data.imageUrl);
          } else {
            setImageUrl('#');
          }
        })
        .catch(error => {
          console.error('Error fetching image URL:', error);
          setImageUrl('#');
        })
        .finally(() => {
          setImageLoading(false);
        });
    } else {
      setImageUrl('#');
      setImageLoading(false);
    }
  }, [currentImagePath, bucketName, currentIndex]); // Added currentIndex to dependencies

  if (imagePaths.length === 0) {
    return (
      <div className="bg-gray-100 rounded-lg p-8 h-96 flex items-center justify-center">
        <div className="text-center text-gray-500">
          <ImageIcon className="w-16 h-16 mx-auto mb-2" />
          <p className="text-sm">No images available</p>
        </div>
      </div>
    );
  }

  return (
    <div>
      {/* Navigation */}
      <div className="flex items-center justify-between mb-3">
        <button
          onClick={() => onIndexChange(Math.max(0, currentIndex - 1))}
          disabled={currentIndex === 0}
          className="flex items-center px-3 py-2 bg-blue-600 text-white rounded-lg text-sm disabled:bg-gray-300 disabled:cursor-not-allowed hover:bg-blue-700 transition-colors"
        >
          ← Previous
        </button>
        <span className="text-sm text-gray-600 font-medium">
          {currentIndex + 1} of {imagePaths.length}
        </span>
        <button
          onClick={() => onIndexChange(Math.min(imagePaths.length - 1, currentIndex + 1))}
          disabled={currentIndex === imagePaths.length - 1}
          className="flex items-center px-3 py-2 bg-blue-600 text-white rounded-lg text-sm disabled:bg-gray-300 disabled:cursor-not-allowed hover:bg-blue-700 transition-colors"
        >
          Next →
        </button>
      </div>
      
      {/* Image */}
      <div className="bg-gray-100 rounded-lg p-4 h-96 flex items-center justify-center mb-4 relative">
        {imageLoading && (
          <div className="absolute inset-0 flex items-center justify-center bg-gray-100 rounded-lg z-10">
            <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
          </div>
        )}
        {imageUrl !== '#' ? (
          <img 
            src={imageUrl} 
            alt={`Image ${currentIndex + 1}`}
            className="max-w-full max-h-full object-contain rounded shadow-lg"
            onLoad={() => setImageLoading(false)}
            onError={(e) => {
              setImageLoading(false);
              console.error('Image failed to load:', imageUrl);
              e.currentTarget.style.display = 'none';
              e.currentTarget.nextElementSibling?.classList.remove('hidden');
            }}
          />
        ) : null}
        <div className={`text-center text-gray-500 ${imageUrl === '#' ? '' : 'hidden'}`}>
          <ImageIcon className="w-16 h-16 mx-auto mb-2" />
          <p className="text-sm">Image not found: {currentImagePath}</p>
        </div>
      </div>
      
      {/* Image Info */}
      <div className="text-center text-sm text-gray-600 mb-4">
        <p className="font-medium">Human: {rowData.human}</p>
        <p className="font-medium">AI: {rowData.ai}</p>
        <p className="text-xs text-gray-500 mt-1">{currentImagePath}</p>
      </div>
    </div>
  );
};

function DashboardPageInner() {
  const params = useParams();
  const router = useRouter();
  const searchParams = useSearchParams();
  const backPath = searchParams.get('source') === 'admin' ? '/admin' : '/batches';
  const backLabel = searchParams.get('source') === 'admin' ? 'Back to admin' : 'Back to batches';
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [selectedRow, setSelectedRow] = useState<number | null>(null);
  const [loading, setLoading] = useState(true);
  const [showImageSidebar, setShowImageSidebar] = useState(false);
  const [currentImageIndex, setCurrentImageIndex] = useState<number>(0);
  const [isExporting, setIsExporting] = useState(false);
  const [sheetImageTags, setSheetImageTags] = useState<{ [imagePath: string]: string[] }>({});
  // Ref to track current sheetImageTags for synchronous access (React 18 batches state updates)
  const sheetImageTagsRef = useRef<{ [imagePath: string]: string[] }>({});
  const [globalImageSpecies, setGlobalImageSpecies] = useState<{ [imagePath: string]: string }>({});
  // Ref to track current globalImageSpecies for synchronous access (React 18 batches state updates)
  const globalImageSpeciesRef = useRef<{ [imagePath: string]: string }>({});
  const [showDownloadOptions, setShowDownloadOptions] = useState(false);
  const [downloadOptions, setDownloadOptions] = useState({
    taggedExcel: true,
    updatedCsv: true,
    originalData: false
  });
  // Track saving state for tags (imagePath -> tag -> saving)
  const [tagSavingState, setTagSavingState] = useState<{ [key: string]: boolean }>({});
  // Track saving state for species classification (imagePath -> saving)
  const [speciesSavingState, setSpeciesSavingState] = useState<{ [imagePath: string]: boolean }>({});
  // Notification state
  const [notification, setNotification] = useState<{ message: string; type: 'success' | 'error' } | null>(null);

  // Keep sheetImageTagsRef in sync with state (for synchronous access in callbacks)
  useEffect(() => {
    sheetImageTagsRef.current = sheetImageTags;
  }, [sheetImageTags]);

  // Fetch sheet-specific image tags
  const fetchSheetTags = useCallback(async (filename: string, sheetName: string) => {
    try {
      const response = await fetch(`/api/global-tags?filename=${encodeURIComponent(filename)}&sheetName=${encodeURIComponent(sheetName)}`, { cache: 'no-store', credentials: 'include' });
      if (response.ok) {
        const data = await response.json();
        const tags = data.sheetImageTags || {};
        setSheetImageTags(tags);
        sheetImageTagsRef.current = tags; // Also update ref immediately
      } else {
        console.error('Failed to fetch sheet-specific tags:', response.status);
      }
    } catch (error) {
      console.error('Error fetching sheet-specific tags:', error);
    }
  }, []);

  // Keep globalImageSpeciesRef in sync with state
  useEffect(() => {
    globalImageSpeciesRef.current = globalImageSpecies;
  }, [globalImageSpecies]);

  // Fetch global image species
  const fetchGlobalSpecies = useCallback(async (filename: string) => {
    try {
      const response = await fetch(`/api/global-species?filename=${encodeURIComponent(filename)}`, { cache: 'no-store', credentials: 'include' });
      if (response.ok) {
        const data = await response.json();
        const species = data.globalImageSpecies || {};
        setGlobalImageSpecies(species);
        globalImageSpeciesRef.current = species; // Also update ref immediately
      }
    } catch (error) {
      console.error('Error fetching global species:', error);
    }
  }, []);

  // Fetch initial data and sheets
  const fetchData = useCallback(async () => {
    try {
      const response = await fetch(`/api/data/${params.id}`, {
        credentials: 'include',
      });
      if (!response.ok) {
        if (response.status === 401 || response.status === 403) {
          router.replace(backPath);
        }
        return;
      }
      const data = await response.json();

      // Fetch global species (no need to fetch tags here yet, we'll fetch when sheet is selected)
      await fetchGlobalSpecies(data.filename);

      const encodedFilename = encodeURIComponent(data.filename);
      const sheetsResponse = await fetch(
        `/api/sheets/${encodedFilename}?batchId=${params.id}`,
        { credentials: 'include' }
      );
      if (!sheetsResponse.ok) {
        console.error('Failed to load sheets:', sheetsResponse.status);
        return;
      }
      const sheetsData = await sheetsResponse.json();
      
      // Remove duplicate sheet names
      const uniqueSheets = sheetsData.filter((sheet: Sheet, index: number, self: Sheet[]) => 
        index === self.findIndex(s => s.sheetName === sheet.sheetName)
      );
      
      setSheets(uniqueSheets);
      setExcelData(null);
      setSelectedSheet('');
    } catch (error) {
      console.error('Error fetching data:', error);
    } finally {
      setLoading(false);
    }
  }, [params.id, fetchGlobalSpecies, router, backPath]);

  // Handle sheet change
  const handleSheetChange = useCallback(async (sheetId: string) => {
    try {
      // Find the sheet name from the sheets array
      const selectedSheetData = sheets.find(sheet => sheet.id === sheetId);
      if (!selectedSheetData) {
        console.error('Sheet not found:', sheetId);
        return;
      }
      
      // Update selected sheet immediately
      setSelectedSheet(selectedSheetData.sheetName);
      
      const response = await fetch(`/api/data/${sheetId}`, {
        credentials: 'include',
      });
      if (!response.ok) return;
      const data = await response.json();
      
      setExcelData(data);
      
      // Fetch sheet-specific tags and global species when switching sheets
      await fetchSheetTags(data.filename, selectedSheetData.sheetName);
      await fetchGlobalSpecies(data.filename);
      
      setSelectedRow(null);
      setShowImageSidebar(false);
    } catch (error) {
      console.error('Error fetching sheet data:', error);
    }
  }, [sheets, fetchSheetTags, fetchGlobalSpecies]);

  // Handle row selection
  const handleRowSelect = useCallback((index: number) => {
    if (selectedRow === index) {
      setSelectedRow(null);
      setShowImageSidebar(false);
    } else {
      setSelectedRow(index);
      setShowImageSidebar(true);
      setCurrentImageIndex(0);
    }
  }, [selectedRow]);

  // Handle tag selection for individual images using sheet-specific tags
  const handleTagSelect = useCallback(async (rowIndex: number, tag: string, imagePath: string) => {
    if (!excelData) return;

    // Create a unique key for this tag operation
    const savingKey = `${imagePath}-${tag}`;
    
    // Set loading state
    setTagSavingState(prev => ({ ...prev, [savingKey]: true }));

    // Calculate the new tags BEFORE state update using ref (React 18 batches state updates)
    const previousImageTags = sheetImageTagsRef.current[imagePath] || [];
    let newTags: string[];
    
    // Toggle the tag for this specific image
    if (previousImageTags.includes(tag)) {
      newTags = previousImageTags.filter(t => t !== tag);
    } else {
      newTags = [...previousImageTags, tag];
    }
    
    // Update state and ref
    const updatedTags = { ...sheetImageTagsRef.current, [imagePath]: newTags };
    sheetImageTagsRef.current = updatedTags;
    setSheetImageTags(updatedTags);
    
    // Helper function to rollback
    const rollback = () => {
      const rolledBackTags = { ...sheetImageTagsRef.current, [imagePath]: previousImageTags };
      sheetImageTagsRef.current = rolledBackTags;
      setSheetImageTags(rolledBackTags);
    };

    try {
      // Update sheet-specific tags on the server (timeout to avoid hanging when Vercel returns "---")
      const response = await fetchWithTimeout('/api/global-tags', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          filename: excelData.filename,
          sheetName: excelData.sheetName,
          imagePath,
          tags: newTags
        })
      });
      
      // Check if response is successful before parsing JSON
      if (!response.ok) {
        const errorText = await response.text();
        let errorData;
        try {
          errorData = JSON.parse(errorText);
        } catch {
          errorData = { error: errorText || 'Unknown error' };
        }
        console.error('Failed to update sheet-specific tags:', errorData);
        
        // Rollback optimistic update on failure
        rollback();
        
        // Show error notification
        setNotification({ 
          message: `Failed to save tag "${tag}": ${errorData.error || 'Network error'}`, 
          type: 'error' 
        });
        setTimeout(() => setNotification(null), 5000);
        return;
      }
      
      const result = await response.json();
      if (!result.success) {
        console.error('Failed to update sheet-specific tags:', result);
        
        // Rollback optimistic update on failure
        rollback();
        
        // Show error notification
        setNotification({ 
          message: `Failed to save tag "${tag}": ${result.error || 'Unknown error'}`, 
          type: 'error' 
        });
        setTimeout(() => setNotification(null), 5000);
      } else {
        // Show success notification
        setNotification({ 
          message: `Tag "${tag}" saved successfully`, 
          type: 'success' 
        });
        setTimeout(() => setNotification(null), 3000);
      }
    } catch (error) {
      console.error('Error updating sheet-specific tags:', error);
      const isTimeout = error instanceof Error && error.name === 'AbortError';
      
      // Rollback optimistic update on network error
      rollback();
      
      // Show error notification (include timeout message when Vercel returns "---")
      setNotification({ 
        message: `Failed to save tag "${tag}": ${isTimeout ? 'Request timed out. Please try again.' : error instanceof Error ? error.message : 'Network error'}`, 
        type: 'error' 
      });
      setTimeout(() => setNotification(null), 5000);
    } finally {
      // Clear loading state
      setTagSavingState(prev => {
        const newState = { ...prev };
        delete newState[savingKey];
        return newState;
      });
    }
  }, [excelData]);

  // Handle species classification
  const handleSpeciesClassification = useCallback(async (rowIndex: number, species: string, imagePath: string) => {
    if (!excelData) return;

    // Set loading state
    setSpeciesSavingState(prev => ({ ...prev, [imagePath]: true }));

    // Capture previous value for rollback using ref (React 18 batches state updates)
    const previousSpecies = globalImageSpeciesRef.current[imagePath];

    // Helper function to rollback species
    const rollbackSpecies = () => {
      if (previousSpecies !== undefined) {
        const rolledBack = { ...globalImageSpeciesRef.current, [imagePath]: previousSpecies };
        globalImageSpeciesRef.current = rolledBack;
        setGlobalImageSpecies(rolledBack);
      } else {
        // Previous was undefined, so just delete the entry
        const rolledBack = { ...globalImageSpeciesRef.current };
        delete rolledBack[imagePath];
        globalImageSpeciesRef.current = rolledBack;
        setGlobalImageSpecies(rolledBack);
      }
    };

    // Handle clear selection
    if (species === 'CLEAR_SELECTION') {
      // Update UI and ref optimistically
      const updatedSpecies = { ...globalImageSpeciesRef.current };
      delete updatedSpecies[imagePath];
      globalImageSpeciesRef.current = updatedSpecies;
      setGlobalImageSpecies(updatedSpecies);
      
      // Clear from database
      try {
        const response = await fetchWithTimeout('/api/global-species', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            filename: excelData.filename,
            imagePath,
            species: '' // Empty string to clear
          })
        });

        if (!response.ok) {
          const errorText = await response.text();
          let errorData;
          try {
            errorData = JSON.parse(errorText);
          } catch {
            errorData = { error: errorText || 'Unknown error' };
          }
          
          // Rollback optimistic update
          rollbackSpecies();
          
          setNotification({ 
            message: `Failed to clear species: ${errorData.error || 'Network error'}`, 
            type: 'error' 
          });
          setTimeout(() => setNotification(null), 5000);
          return;
        }

        const result = await response.json();
        if (!result.success) {
          // Rollback optimistic update
          rollbackSpecies();
          
          setNotification({ 
            message: `Failed to clear species: ${result.error || 'Unknown error'}`, 
            type: 'error' 
          });
          setTimeout(() => setNotification(null), 5000);
        } else {
          setNotification({ 
            message: 'Species classification cleared successfully', 
            type: 'success' 
          });
          setTimeout(() => setNotification(null), 3000);
        }
      } catch (error) {
        console.error('Error clearing species classification:', error);
        
        // Rollback optimistic update
        rollbackSpecies();
        
        setNotification({ 
          message: `Failed to clear species: ${error instanceof Error ? error.message : 'Network error'}`, 
          type: 'error' 
        });
        setTimeout(() => setNotification(null), 5000);
      } finally {
        setSpeciesSavingState(prev => {
          const newState = { ...prev };
          delete newState[imagePath];
          return newState;
        });
      }
      return;
    }

    // Update local state and ref immediately
    const updatedSpecies = { ...globalImageSpeciesRef.current, [imagePath]: species };
    globalImageSpeciesRef.current = updatedSpecies;
    setGlobalImageSpecies(updatedSpecies);

    try {
      // Update species classification on the server (CSV)
      const csvResponse = await fetchWithTimeout('/api/update-species', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          excelSheetId: excelData._id,
          filename: excelData.filename,
          sheetName: excelData.sheetName,
          imagePath,
          species
        })
      });
      
      // Update global species in database
      const globalResponse = await fetchWithTimeout('/api/global-species', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          filename: excelData.filename,
          imagePath,
          species
        })
      });
      
      // Check CSV response
      if (!csvResponse.ok) {
        const errorText = await csvResponse.text();
        let errorData;
        try {
          errorData = JSON.parse(errorText);
        } catch {
          errorData = { error: errorText || 'Unknown error' };
        }
        
        // Rollback optimistic update
        rollbackSpecies();
        
        setNotification({ 
          message: `Failed to save species: ${errorData.error || 'Network error'}`, 
          type: 'error' 
        });
        setTimeout(() => setNotification(null), 5000);
        return;
      }

      // Check global response
      if (!globalResponse.ok) {
        const errorText = await globalResponse.text();
        let errorData;
        try {
          errorData = JSON.parse(errorText);
        } catch {
          errorData = { error: errorText || 'Unknown error' };
        }
        
        // Rollback optimistic update
        rollbackSpecies();
        
        setNotification({ 
          message: `Failed to save species: ${errorData.error || 'Network error'}`, 
          type: 'error' 
        });
        setTimeout(() => setNotification(null), 5000);
        return;
      }
      
      const csvResult = await csvResponse.json();
      const globalResult = await globalResponse.json();
      
      if (!csvResult.success || !globalResult.success) {
        // Rollback optimistic update
        rollbackSpecies();
        
        setNotification({ 
          message: `Failed to save species: ${csvResult.error || globalResult.error || 'Unknown error'}`, 
          type: 'error' 
        });
        setTimeout(() => setNotification(null), 5000);
      } else {
        setNotification({ 
          message: `Species "${species}" saved successfully`, 
          type: 'success' 
        });
        setTimeout(() => setNotification(null), 3000);
      }
    } catch (error) {
      console.error('Error updating species classification:', error);
      const isTimeout = error instanceof Error && error.name === 'AbortError';
      
      // Rollback optimistic update
      rollbackSpecies();
      
      setNotification({ 
        message: `Failed to save species: ${isTimeout ? 'Request timed out. Please try again.' : error instanceof Error ? error.message : 'Network error'}`, 
        type: 'error' 
      });
      setTimeout(() => setNotification(null), 5000);
    } finally {
      // Clear loading state
      setSpeciesSavingState(prev => {
        const newState = { ...prev };
        delete newState[imagePath];
        return newState;
      });
    }
  }, [excelData]);

  // Handle export tagged data (show download options)
  const handleExportTaggedData = useCallback(async () => {
    if (!excelData) return;
    
    setShowDownloadOptions(true);
  }, [excelData]);

  // Handle download with selected options
  const handleBackToBatches = useCallback(() => {
    router.push(backPath);
  }, [router, backPath]);

  const handleDownloadWithOptions = useCallback(async () => {
    if (!excelData) return;
    
    setIsExporting(true);
    setShowDownloadOptions(false);
    
    try {
      const downloads: string[] = [];
      const failures: string[] = [];

      const readApiError = async (response: Response) => {
        const ct = response.headers.get('Content-Type') || '';
        if (ct.includes('application/json')) {
          const data = await response.json().catch(() => ({}));
          return (data as { error?: string }).error || response.statusText;
        }
        const text = await response.text();
        return text.slice(0, 200) || response.statusText;
      };
      
      // Download tagged Excel if selected
      if (downloadOptions.taggedExcel) {
        const response = await fetch('/api/export-tags', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          credentials: 'include',
          body: JSON.stringify({
            filename: excelData.filename,
            bucketName: excelData.bucketName,
          }),
        });
        
        if (response.ok) {
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `tagged_data_${excelData.filename.replace('.xlsx', '')}.xlsx`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          window.URL.revokeObjectURL(url);
          downloads.push('Tagged Excel');
        } else {
          failures.push(`Tagged Excel: ${await readApiError(response)}`);
        }
      }
      
      // Download updated CSV if selected
      if (downloadOptions.updatedCsv) {
        const response = await fetch(
          `/api/download-csv?batchId=${encodeURIComponent(String(excelData._id))}`,
          { credentials: 'include' }
        );
        
        if (response.ok) {
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          const cd = response.headers.get('Content-Disposition');
          const filenameMatch = cd?.match(/filename="([^"]+)"/);
          a.download = filenameMatch?.[1] ?? 'updated_species_data.csv';
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          window.URL.revokeObjectURL(url);
          downloads.push('Updated CSV');
        } else {
          failures.push(`Updated CSV: ${await readApiError(response)}`);
        }
      }
      
      // Download original data if selected
      if (downloadOptions.originalData) {
        const response = await fetch(`/api/download-original/${excelData._id}`, {
          credentials: 'include',
        });
        
        if (response.ok) {
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `original_${excelData.filename}`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          window.URL.revokeObjectURL(url);
          downloads.push('Original Data');
        } else {
          failures.push(`Original Excel: ${await readApiError(response)}`);
        }
      }
      
      const wanted =
        downloadOptions.taggedExcel ||
        downloadOptions.updatedCsv ||
        downloadOptions.originalData;

      if (!wanted) {
        alert('No files selected for download');
      } else if (downloads.length > 0 && failures.length === 0) {
        alert(`Downloaded: ${downloads.join(', ')}`);
      } else if (downloads.length > 0 && failures.length > 0) {
        alert(
          `Downloaded: ${downloads.join(', ')}\n\nCould not complete:\n${failures.join('\n')}`
        );
      } else {
        alert(
          failures.length > 0
            ? failures.join('\n\n')
            : 'Download failed. Please try again.'
        );
      }
      
    } catch (error) {
      console.error('Download failed:', error);
      alert('Download failed. Please try again.');
    } finally {
      setIsExporting(false);
    }
  }, [excelData, downloadOptions]);

  // Get current row data and image paths
  const currentRowData = useMemo(() => {
    if (!excelData || selectedRow === null || !excelData.data || selectedRow >= excelData.data.length) return null;
    const row = excelData.data[selectedRow];
    if (!row) return null; // Add null check for row
    
    const imagePaths = getImagePaths(row);
    
    return {
      row,
      imagePaths
    };
  }, [excelData, selectedRow]);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  // Notification component
  const NotificationToast = notification && (
    <div className={`fixed top-20 right-4 z-50 p-4 rounded-lg shadow-lg max-w-md animate-in slide-in-from-top-5 ${
      notification.type === 'success' 
        ? 'bg-green-50 border border-green-200 text-green-800' 
        : 'bg-red-50 border border-red-200 text-red-800'
    }`}>
      <div className="flex items-center justify-between">
        <div className="flex items-center">
          {notification.type === 'success' ? (
            <svg className="w-5 h-5 mr-2" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
            </svg>
          ) : (
            <svg className="w-5 h-5 mr-2" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
            </svg>
          )}
          <span className="text-sm font-medium">{notification.message}</span>
        </div>
        <button
          onClick={() => setNotification(null)}
          className="ml-4 text-gray-400 hover:text-gray-600"
        >
          <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
            <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
          </svg>
        </button>
      </div>
    </div>
  );

  if (!excelData) {
    return (
      <div className="min-h-screen bg-gray-50">
        <Navbar 
          selectedSheet={selectedSheet} 
          sheets={sheets} 
          onSheetChange={handleSheetChange}
          onExport={handleExportTaggedData}
          isExporting={isExporting}
          onBack={handleBackToBatches}
          backLabel={backLabel}
        />
        <div className="flex items-center justify-center min-h-[calc(100vh-64px)]">
          <div className="text-center">
            <h1 className="text-2xl font-bold text-gray-900 mb-2">Select a Taxonomic Level</h1>
            <p className="text-gray-600">Choose a taxonomic level from the dropdown above to view the data and images.</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Notification Toast */}
      {NotificationToast}
      
      <Navbar 
        selectedSheet={selectedSheet} 
        sheets={sheets} 
        onSheetChange={handleSheetChange}
        onExport={handleExportTaggedData}
        isExporting={isExporting}
        onBack={handleBackToBatches}
        backLabel={backLabel}
      />

      <div className="flex">
        {/* Main Content */}
        <div className={`flex-1 transition-all duration-300 ${showImageSidebar ? 'mr-[800px]' : ''}`}>
          <div className="p-6">
            <div className="bg-white rounded-lg shadow">
              <div className="px-6 py-4 border-b border-gray-200">
                <h2 className="text-lg font-medium text-gray-900">Data Table</h2>
                <p className="text-sm text-gray-600">Select rows to view images and add tags</p>
              </div>
              
              <div className="overflow-x-auto">
                <table className="min-w-full divide-y divide-gray-200">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Select
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Human
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        AI
                      </th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {excelData.data.map((row, index) => (
                      <tr 
                        key={index} 
                        className={`${selectedRow === index ? 'bg-blue-50' : 'hover:bg-gray-50'} cursor-pointer`} 
                        onClick={() => handleRowSelect(index)}
                      >
                        <td className="px-6 py-4 whitespace-nowrap">
                          <input
                            type="radio"
                            checked={selectedRow === index}
                            onChange={() => handleRowSelect(index)}
                            className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300"
                          />
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                          {row.human}
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                          {row.ai}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>

        {/* Image Sidebar */}
        {showImageSidebar && currentRowData && (
          <div className="fixed right-0 top-16 bottom-0 w-[800px] bg-white shadow-lg border-l border-gray-200 overflow-y-auto">
            <div className="p-6">
              <div className="flex items-center justify-between mb-6">
                <h3 className="text-lg font-medium text-gray-900 flex items-center">
                  <ImageIcon className="w-5 h-5 mr-2" />
                  Images & Tags
                </h3>
                <button
                  onClick={() => setShowImageSidebar(false)}
                  className="text-gray-400 hover:text-gray-600"
                >
                  ×
                </button>
              </div>

              {/* Image List */}
              <div className="mb-6">
                <h4 className="text-sm font-medium text-gray-900 mb-3">Row {selectedRow! + 1}</h4>
                <div className="border border-gray-200 rounded-lg">
                  <div className="flex items-center p-3 bg-blue-50 rounded-t-lg">
                    <ImageIcon className="w-5 h-5 mr-2 text-blue-600" />
                    <span className="text-sm text-gray-700 font-medium">
                      {currentRowData.imagePaths.length > 0 
                        ? `${currentRowData.imagePaths.length} images` 
                        : 'No images'
                      }
                    </span>
                    <span className="text-xs text-gray-500 ml-auto">
                      {Object.keys(sheetImageTags).filter(imagePath => 
                        currentRowData.imagePaths.includes(imagePath) && 
                        sheetImageTags[imagePath].length > 0
                      ).length} images tagged
                    </span>
                  </div>
                  
                  {currentRowData.imagePaths.length > 0 ? (
                    <div className="p-3">
                      <div className="grid grid-cols-1 gap-1 max-h-32 overflow-y-auto">
                        {currentRowData.imagePaths.map((filename, idx) => {
                          const hasImageTags = (sheetImageTags[filename]?.length || 0) > 0;
                          const hasSpeciesClassification = globalImageSpecies[filename];
                          return (
                            <div 
                              key={idx} 
                              className={`flex items-center p-2 rounded text-xs transition-colors cursor-pointer ${
                                idx === currentImageIndex ? 'bg-blue-100 text-blue-700' : 'bg-gray-50 text-gray-700 hover:bg-gray-100'
                              }`}
                              onClick={() => setCurrentImageIndex(idx)}
                            >
                              <ImageIcon className="w-3 h-3 mr-2 text-gray-400" />
                              <span className="truncate flex-1">{filename}</span>
                              <div className="flex items-center ml-1 space-x-1">
                                {hasImageTags && (
                                  <div className="flex items-center">
                                    <TagIcon className="w-3 h-3 text-green-500" />
                                    <span className="text-green-600 font-medium ml-1">
                                      {sheetImageTags[filename]?.length || 0}
                                    </span>
                                  </div>
                                )}
                                {hasSpeciesClassification && (
                                  <div className="flex items-center">
                                    <span className="text-blue-500 font-medium">🐾</span>
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ) : (
                    <div className="p-3">
                      <div className="text-xs text-gray-500">No images found for this row</div>
                    </div>
                  )}
                </div>
              </div>

              {/* Image Display */}
              <div className="mb-6">
                <ImageDisplay
                  imagePaths={currentRowData.imagePaths}
                  currentIndex={currentImageIndex}
                  onIndexChange={setCurrentImageIndex}
                  rowData={{ human: currentRowData.row.human, ai: currentRowData.row.ai }}
                  bucketName={excelData?.bucketName || ''}
                />
              </div>

              {/* Tags Section */}
              <div>
                <h5 className="text-sm font-medium text-gray-900 mb-3 flex items-center">
                  <TagIcon className="w-4 h-4 mr-1" />
                  Tags for Current Image
                </h5>
                <div className="mb-3 p-2 bg-blue-50 rounded text-xs text-blue-700">
                  Tagging: {currentRowData.imagePaths[currentImageIndex]}
                </div>
                <div className="space-y-2">
                  {TAGS.map((tag) => {
                    const currentImagePath = currentRowData.imagePaths[currentImageIndex];
                    const isChecked = sheetImageTags[currentImagePath]?.includes(tag) || false;
                    const savingKey = `${currentImagePath}-${tag}`;
                    const isSaving = tagSavingState[savingKey] || false;
                    
                    return (
                      <label key={tag} className="flex items-center relative">
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => handleTagSelect(selectedRow!, tag, currentImagePath)}
                          disabled={isSaving}
                          className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mr-2 disabled:opacity-50 disabled:cursor-not-allowed"
                        />
                        <span className="text-sm text-gray-700 flex-1">{tag}</span>
                        {isSaving && (
                          <span className="ml-2 text-xs text-blue-600 flex items-center">
                            <span className="animate-spin rounded-full h-3 w-3 border-b-2 border-blue-600 mr-1"></span>
                            Saving...
                          </span>
                        )}
                      </label>
                    );
                  })}
                </div>
              </div>

              {/* Species Classification Section */}
              <div className="mt-6">
                <h5 className="text-sm font-medium text-gray-900 mb-3 flex items-center">
                  🐾 Species Classification
                </h5>
                <div className="mb-3 p-2 bg-green-50 rounded text-xs text-green-700">
                  Update species for: {currentRowData.imagePaths[currentImageIndex]}
                </div>
                <div className="relative">
                  <select 
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 disabled:opacity-50 disabled:cursor-not-allowed disabled:bg-gray-100"
                    value={globalImageSpecies[currentRowData.imagePaths[currentImageIndex]] || ""}
                    onChange={(e) => handleSpeciesClassification(selectedRow!, e.target.value, currentRowData.imagePaths[currentImageIndex])}
                    disabled={speciesSavingState[currentRowData.imagePaths[currentImageIndex]] || false}
                  >
                    <option value="" disabled>Select Species Classification</option>
                    <option value="CLEAR_SELECTION" className="text-red-600 font-medium">
                      ✕ Clear Selection
                    </option>
                    {SPECIES_LIST.map((species) => (
                      <option key={species} value={species}>
                        {species}
                      </option>
                    ))}
                  </select>
                  {speciesSavingState[currentRowData.imagePaths[currentImageIndex]] && (
                    <div className="absolute right-3 top-1/2 transform -translate-y-1/2 flex items-center text-blue-600">
                      <span className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600 mr-2"></span>
                      <span className="text-xs">Saving...</span>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Download Options Modal */}
      {showDownloadOptions && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-full max-w-md">
            <h3 className="text-lg font-medium text-gray-900 mb-4">Select Files to Download</h3>
            
            <div className="space-y-3 mb-6">
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={downloadOptions.taggedExcel}
                  onChange={(e) => setDownloadOptions(prev => ({ ...prev, taggedExcel: e.target.checked }))}
                  className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mr-3"
                />
                <span className="text-sm text-gray-700">Tagged Excel Analysis</span>
              </label>
              
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={downloadOptions.updatedCsv}
                  onChange={(e) => setDownloadOptions(prev => ({ ...prev, updatedCsv: e.target.checked }))}
                  className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mr-3"
                />
                <span className="text-sm text-gray-700">Updated Species CSV</span>
              </label>
              
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={downloadOptions.originalData}
                  onChange={(e) => setDownloadOptions(prev => ({ ...prev, originalData: e.target.checked }))}
                  className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mr-3"
                />
                <span className="text-sm text-gray-700">Original Excel Data</span>
              </label>
            </div>
            
            <div className="flex space-x-3">
              <button
                onClick={() => setShowDownloadOptions(false)}
                className="flex-1 px-4 py-2 border border-gray-300 rounded-md text-sm font-medium text-gray-700 hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                onClick={handleDownloadWithOptions}
                disabled={!downloadOptions.taggedExcel && !downloadOptions.updatedCsv && !downloadOptions.originalData}
                className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-medium hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed"
              >
                Download Selected
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default function DashboardPage() {
  return (
    <Suspense
      fallback={
        <div className="min-h-screen bg-gray-50 flex items-center justify-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600" />
        </div>
      }
    >
      <DashboardPageInner />
    </Suspense>
  );
}