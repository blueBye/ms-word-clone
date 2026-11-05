/** @type {import('next').NextConfig} */
const config = {
  // Cloudflare Pages serves prebuilt assets. The DOCX engine and editor run in
  // the browser, so no Node.js runtime or Cloudflare Workers are required.
  output: 'export',
  trailingSlash: true,
  images: { unoptimized: true },

  // The app consumes the root workspace package after `build:engine`.
  transpilePackages: ['@local/docx-ir-engine'],
};

export default config;
