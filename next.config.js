/** @type {import('next').NextConfig} */
const nextConfig = {
  images: {
    remotePatterns: [
      {
        protocol: 'https',
        hostname: 'graph.microsoft.com',
        port: '',
        pathname: '/v1.0/me/drive/items/**',
      },
    ],
  },
}

module.exports = nextConfig
