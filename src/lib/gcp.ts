import { Storage } from '@google-cloud/storage';

// Initialize GCP Storage client with individual environment variables
const getStorageClient = () => {
  try {
    // Check if we have individual credential environment variables
    if (process.env.GOOGLE_CLIENT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
      const credentials = {
        type: "service_account",
        project_id: process.env.GOOGLE_CLOUD_PROJECT_ID,
        private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
        private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        client_id: process.env.GOOGLE_CLIENT_ID,
        auth_uri: "https://accounts.google.com/o/oauth2/auth",
        token_uri: "https://oauth2.googleapis.com/token",
        auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
        client_x509_cert_url: `https://www.googleapis.com/robot/v1/metadata/x509/${encodeURIComponent(process.env.GOOGLE_CLIENT_EMAIL)}`,
        universe_domain: "googleapis.com"
      };
      
      return new Storage({
        projectId: process.env.GOOGLE_CLOUD_PROJECT_ID,
        credentials,
      });
    } else if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
      // Fallback to JSON string approach
      const credentials = JSON.parse(process.env.GOOGLE_APPLICATION_CREDENTIALS);
      
      return new Storage({
        projectId: process.env.GOOGLE_CLOUD_PROJECT_ID,
        credentials,
      });
    } else {
      // Fallback to default authentication (useful for local development)
      return new Storage({
        projectId: process.env.GOOGLE_CLOUD_PROJECT_ID,
      });
    }
  } catch (error) {
    console.error('Error initializing GCP Storage client:', error);
    throw new Error('Failed to initialize GCP Storage client. Check your Google Cloud credentials environment variables.');
  }
};

let cachedStorage: Storage | null = null;
function getStorage(): Storage {
  if (!cachedStorage) cachedStorage = getStorageClient();
  return cachedStorage;
}

function requireMasterSheetUploadBucket(): string {
  const bucket =
    process.env.MASTER_SHEET_UPLOAD_BUCKET || process.env.GCP_DEFAULT_BUCKET;
  if (!bucket) {
    throw new Error(
      'MASTER_SHEET_UPLOAD_BUCKET or GCP_DEFAULT_BUCKET is required for master-sheet uploads.'
    );
  }
  return bucket;
}

export interface BucketInfo {
  name: string;
  location: string;
  created: Date;
}

export interface ImageInfo {
  name: string;
  url: string;
  size: number;
  contentType: string;
}

/**
 * List all available GCP buckets
 */
export async function listBuckets(): Promise<BucketInfo[]> {
  try {
    const storage = getStorage();
    const [buckets] = await storage.getBuckets();
    return buckets.map(bucket => ({
      name: bucket.name,
      location: bucket.metadata?.location || 'Unknown',
      created: bucket.metadata?.timeCreated ? new Date(bucket.metadata.timeCreated) : new Date()
    }));
  } catch (error) {
    console.error('Error listing buckets:', error);
    throw new Error('Failed to list GCP buckets');
  }
}

/**
 * List images in a specific bucket
 */
export async function listImagesInBucket(bucketName: string, prefix?: string): Promise<ImageInfo[]> {
  try {
    const storage = getStorage();
    const bucket = storage.bucket(bucketName);
    const [files] = await bucket.getFiles({ prefix });
    
    return files.map(file => ({
      name: file.name,
      url: `https://storage.googleapis.com/${bucketName}/${file.name}`,
      size: parseInt(String(file.metadata?.size || '0')),
      contentType: file.metadata?.contentType || 'image/jpeg'
    }));
  } catch (error) {
    console.error('Error listing images in bucket:', error);
    throw new Error(`Failed to list images in bucket: ${bucketName}`);
  }
}

/**
 * Get a signed URL for an image in a bucket
 */
export async function getSignedUrl(bucketName: string, fileName: string, expiresIn: number = 3600): Promise<string> {
  try {
    const storage = getStorage();
    const bucket = storage.bucket(bucketName);
    const file = bucket.file(fileName);
    
    const [signedUrl] = await file.getSignedUrl({
      action: 'read',
      expires: Date.now() + expiresIn * 1000,
    });
    
    return signedUrl;
  } catch (error) {
    console.error('Error generating signed URL:', error);
    throw new Error(`Failed to generate signed URL for ${fileName}`);
  }
}

/**
 * Check if a bucket exists and is accessible
 */
export async function checkBucketAccess(bucketName: string): Promise<boolean> {
  try {
    const storage = getStorage();
    const bucket = storage.bucket(bucketName);
    const [exists] = await bucket.exists();
    return exists;
  } catch (error) {
    console.error('Error checking bucket access:', error);
    return false;
  }
}

function sanitizeObjectName(name: string): string {
  return name.replace(/[^a-zA-Z0-9._-]/g, "_");
}

export async function createMasterSheetUploadUrl(params: {
  objectPath: string;
  contentType?: string;
  expiresInSeconds?: number;
}): Promise<{ bucketName: string; uploadUrl: string; objectPath: string }> {
  const bucketName = requireMasterSheetUploadBucket();
  const storage = getStorage();
  const file = storage.bucket(bucketName).file(params.objectPath);
  const [uploadUrl] = await file.getSignedUrl({
    version: "v4",
    action: "write",
    expires: Date.now() + (params.expiresInSeconds ?? 900) * 1000,
    contentType: params.contentType || "application/octet-stream",
  });
  return { bucketName, uploadUrl, objectPath: params.objectPath };
}

export function buildMasterSheetObjectPath(params: {
  jobId: string;
  field: string;
  originalName: string;
}): string {
  return `master-sheet-inputs/${params.jobId}/${params.field}/${sanitizeObjectName(
    params.originalName || "file"
  )}`;
}

export async function downloadMasterSheetInput(objectPath: string): Promise<Buffer> {
  const bucketName = requireMasterSheetUploadBucket();
  const storage = getStorage();
  const [buffer] = await storage.bucket(bucketName).file(objectPath).download();
  return buffer;
}

export async function deleteMasterSheetInputs(objectPaths: string[]): Promise<void> {
  const bucketName = requireMasterSheetUploadBucket();
  const storage = getStorage();
  await Promise.all(
    objectPaths.map(async (path) => {
      try {
        await storage.bucket(bucketName).file(path).delete({ ignoreNotFound: true });
      } catch {
        // Best-effort cleanup
      }
    })
  );
}
