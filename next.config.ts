import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  serverExternalPackages: ['pizzip'],
  // 로컬 네트워크 IP로 접속할 때 HMR WebSocket을 허용
  allowedDevOrigins: ['192.168.20.232'],
};

export default nextConfig;
