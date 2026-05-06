import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  experimental: {
    // Applies to Route Handler / API uploads proxied through Next server.
    proxyClientMaxBodySize: 700 * 1024 * 1024, // 700 MB
    serverActions: {
      bodySizeLimit: "700mb",
    },
  },
};

export default nextConfig;
