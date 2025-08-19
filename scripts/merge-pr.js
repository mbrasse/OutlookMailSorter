// scripts/merge-pr.js
// Usage in Actions: sets APP_ID, PRIVATE_KEY, GITHUB_REPOSITORY, PULL_NUMBER
// Merges with 'squash' method.
//
// This script is robust: it normalizes newline escapes in the PRIVATE_KEY,
// obtains an installation access token for the repository, optionally waits
// for mergeability to be determined, and performs a squash merge.

const { App } = require('@octokit/app');
const { Octokit } = require('@octokit/rest');

async function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function main() {
  try {
    const appId = process.env.APP_ID;
    let privateKey = process.env.PRIVATE_KEY;
    const repository = process.env.GITHUB_REPOSITORY;
    const pullNumberEnv = process.env.PULL_NUMBER || '';

    if (!appId) throw new Error('APP_ID environment variable is missing.');
    if (!privateKey) throw new Error('PRIVATE_KEY environment variable is missing.');
    if (!repository) throw new Error('GITHUB_REPOSITORY environment variable is missing.');

    // Normalize newline escapes if secret was pasted with \n sequences
    // Only replace literal backslash-n sequences so we don't accidentally change valid keys
    if (privateKey.includes('\\n')) {
      privateKey = privateKey.replace(/\\n/g, '\n');
    }

    const repoParts = repository.split('/');
    if (repoParts.length !== 2) {
      throw new Error(`GITHUB_REPOSITORY must be in the form "owner/repo". Got: "${repository}"`);
    }
    const [owner, repo] = repoParts;

    const pull_number = Number(pullNumberEnv);
    if (!pull_number || Number.isNaN(pull_number)) {
      throw new Error(`Invalid PULL_NUMBER: "${pullNumberEnv}"`);
    }

    // Initialize GitHub App
    const app = new App({ id: Number(appId), privateKey });

    // 1) Create a JWT for the App
    const jwt = app.getSignedJsonWebToken();

    // 2) Use the app JWT to get the installation id for this repo
    const appOctokit = new Octokit({ auth: jwt });
    const installResp = await appOctokit.request('GET /repos/{owner}/{repo}/installation', {
      owner,
      repo,
    });

    const installationId = installResp && installResp.data && installResp.data.id;
    if (!installationId) {
      throw new Error('Could not find installation id for repository');
    }

    // 3) Create an installation access token
    const tokenResp = await appOctokit.request('POST /app/installations/{installation_id}/access_tokens', {
      installation_id: installationId,
    });
    const installationToken = tokenResp && tokenResp.data && tokenResp.data.token;
    if (!installationToken) throw new Error('Failed to obtain installation access token');

    // 4) Optionally check mergeability / status checks here (recommended)
    const octokit = new Octokit({ auth: installationToken });

    // Get PR details
    let pr = await octokit.pulls.get({ owner, repo, pull_number });
    if (!pr || !pr.data) {
      throw new Error(`Failed to fetch PR #${pull_number} data`);
    }

    // Abort if PR is closed
    if (pr.data.state !== 'open') {
      throw new Error(`PR #${pull_number} is not open (state: ${pr.data.state}). Aborting.`);
    }

    // OPTIONAL: abort if PR is a draft
    if (pr.data.draft) {
      throw new Error('PR is a draft. Aborting merge.');
    }

    // pr.data.mergeable can be null while GitHub determines mergeability.
    // Wait for it to become non-null (with a timeout) to make a better decision.
    const maxChecks = 6;
    const delayMs = 2000;
    let checks = 0;
    while ((pr.data.mergeable === null) && checks < maxChecks) {
      console.log(`PR mergeable state is null. Waiting ${delayMs}ms for GitHub to calculate mergeability... (${checks + 1}/${maxChecks})`);
      await sleep(delayMs);
      pr = await octokit.pulls.get({ owner, repo, pull_number });
      checks += 1;
    }

    // If mergeable is false, don't attempt to merge.
    if (pr.data.mergeable === false) {
      // Provide some context: combined status / checks could be failing.
      const mergeableState = pr.data.mergeable_state || 'unknown';
      throw new Error(`PR is not mergeable (mergeable=false, mergeable_state=${mergeableState}). Aborting.`);
    }

    // 5) Merge the PR with squash
    const mergeMethod = 'squash';
    console.log(`Attempting to squash-merge PR #${pull_number} in ${owner}/${repo}...`);
    const mergeResp = await octokit.pulls.merge({
      owner,
      repo,
      pull_number,
      merge_method: mergeMethod,
    });

    console.log('Merge response:', mergeResp && mergeResp.data ? mergeResp.data : mergeResp);
    if (mergeResp && mergeResp.data && mergeResp.data.merged) {
      console.log(`PR #${pull_number} was squash-merged successfully.`);
      process.exit(0);
    } else {
      console.error('Merge did not complete:', mergeResp && mergeResp.data ? mergeResp.data : mergeResp);
      process.exit(2);
    }
  } catch (err) {
    // Print full error for debugging, and ensure message is printed if it's an Error object
    if (err && err.response && err.response.data) {
      console.error('GitHub API error response:', JSON.stringify(err.response.data, null, 2));
    }
    console.error('Error:', err && err.message ? err.message : err);
    process.exit(1);
  }
}

main();