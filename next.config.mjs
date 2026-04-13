/** @type {import('next').NextConfig} */
const nextConfig = {
  // Provide a build-time fallback so next-auth can construct valid URLs
  // during static prerendering (e.g. /_not-found) even when NEXTAUTH_URL
  // is not explicitly present in the build environment.
  env: {
    NEXTAUTH_URL: process.env.NEXTAUTH_URL || 'https://m365-ca-review.vercel.app',
  },
};

export default nextConfig;
