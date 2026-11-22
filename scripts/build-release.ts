import fs from 'node:fs';
import path from 'node:path';
import { generateStandaloneHtml } from '../src/templates/standaloneHtmlGenerator';
import { generateAppsScriptCode } from '../src/templates/appsScriptTemplate';
import {
  APP_VERSION,
  DEFAULT_GITHUB_REPO,
  RELEASE_NOTES,
  SCRIPT_VERSION,
  normalizeGithubRepo,
} from '../src/version';

const rootDir = process.cwd();
const outDir = path.join(rootDir, 'releases');
fs.mkdirSync(outDir, { recursive: true });

// Automatically picks up GITHUB_REPOSITORY when running inside GitHub Actions (e.g. "owner/repo")
const githubRepo = normalizeGithubRepo(
  process.env.GITHUB_REPOSITORY || process.env.M7FAZTY_GITHUB_REPO || DEFAULT_GITHUB_REPO
);

const html = generateStandaloneHtml({
  starterMode: true,
  profileName: '',
  githubRepo,
});

const appsScript = generateAppsScriptCode('');

const versionManifest = {
  version: APP_VERSION,
  scriptVersion: SCRIPT_VERSION,
  updatedAt: new Date().toISOString(),
  githubRepo,
  htmlFile: 'releases/M7FAZTY.html',
  scriptFile: 'releases/Code.gs',
  notes: RELEASE_NOTES,
};

fs.writeFileSync(path.join(outDir, 'M7FAZTY.html'), html, 'utf8');
fs.writeFileSync(path.join(outDir, 'Code.gs'), appsScript, 'utf8');
fs.writeFileSync(path.join(outDir, 'version.json'), JSON.stringify(versionManifest, null, 2) + '\n', 'utf8');

console.log(
  `Built M7FAZTY release v${APP_VERSION} (script v${SCRIPT_VERSION})` +
    (githubRepo ? ` for repo ${githubRepo}` : '') +
    ` -> releases/M7FAZTY.html, releases/Code.gs, releases/version.json`
);
