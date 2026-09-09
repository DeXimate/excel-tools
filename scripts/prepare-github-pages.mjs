import { copyFile, readdir, readFile, writeFile } from 'node:fs/promises';
import { extname, join } from 'node:path';

const outputDirectory = join(process.cwd(), 'dist', 'client');
const repository = process.env.GITHUB_REPO_NAME || 'excel-tools';
const prefix = `/${repository}`;
const textExtensions = new Set(['.html', '.js', '.css', '.json', '.rsc', '.txt', '.xml', '.webmanifest']);

async function rewriteDirectory(directory) {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) await rewriteDirectory(path);
    else if (textExtensions.has(extname(entry.name))) {
      const source = await readFile(path, 'utf8');
      const rewritten = source
        .replaceAll('"/_next/', `"${prefix}/_next/`)
        .replaceAll("'/_next/", `'${prefix}/_next/`)
        .replaceAll('"/favicon.svg"', `"${prefix}/favicon.svg"`)
        .replaceAll("'/favicon.svg'", `'${prefix}/favicon.svg'`);
      if (rewritten !== source) await writeFile(path, rewritten);
    }
  }
}

await rewriteDirectory(outputDirectory);
await writeFile(join(outputDirectory, '.nojekyll'), '');
await copyFile(join(outputDirectory, 'index.html'), join(outputDirectory, '404.html'));

const index = await readFile(join(outputDirectory, 'index.html'), 'utf8');
if (!index.includes(`${prefix}/_next/`) || index.includes('"/_next/')) {
  throw new Error('Les chemins GitHub Pages n’ont pas été préparés correctement.');
}

console.log(`GitHub Pages prêt pour ${prefix}/`);
