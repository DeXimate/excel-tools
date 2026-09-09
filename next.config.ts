import type { NextConfig } from 'next';

const isGithubPages = process.env.GITHUB_PAGES === 'true';
const repoName = process.env.GITHUB_REPO_NAME || '';

const nextConfig: NextConfig = {
	output: 'export',
	assetPrefix: isGithubPages && repoName ? `/${repoName}/` : undefined,
	basePath: isGithubPages && repoName ? `/${repoName}` : undefined,
};

export default nextConfig;
