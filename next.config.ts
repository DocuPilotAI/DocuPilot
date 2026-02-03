import type { NextConfig } from 'next'

const nextConfig: NextConfig = {
  // 开发模式下不显示右下角 N 图标
  devIndicators: false,
  // Enable experimental HTTPS for Office Add-in development
  experimental: {
    // serverActions: true,
  },
  // Set turbopack root to silence warning about multiple lockfiles
  turbopack: {
    root: process.cwd(),
  },
  // Disable strict mode to avoid double mounting issues in Office Add-in
  reactStrictMode: false,
  // Redirects for legacy URLs
  async redirects() {
    return [
      {
        source: '/taskpane.html',
        destination: '/taskpane',
        permanent: true,
      },
    ]
  },
  // Allow Office.js scripts
  async headers() {
    return [
      {
        source: '/:path*',
        headers: [
          {
            key: 'Content-Security-Policy',
            value: "frame-ancestors 'self' https://*.officeapps.live.com https://*.office.com",
          },
          {
            key: 'X-Frame-Options',
            value: 'ALLOWALL',
          },
        ],
      },
    ]
  },
}

export default nextConfig
