// scripts/merge-pr.js
// Usage in Actions: sets APP_ID, PRIVATE_KEY, GITHUB_REPOSITORY, PULL_NUMBER
// Merges with 'squash' method.

const { App } = require('@octokit/app');
const { Octokit } = require('@octokit/rest');

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
    privateKey = privateKey.replace(/\\n/g, '\n');

    const [owner, repo] = repository.split('/');
    const pull_number = Number(pullNumberEnv);
    if (!pull_number || Number.isNaN(pull_number)) {
      throw new Error(`Invalid PULL_NUMBER: "${pullNumberEnv}"`);
    }

    const app = new App({ id: Number(appId), privateKey });

    // 1) Create a JWT for the App
    const jwt = app.getSignedJsonWebToken();

    // 2) Use the app JWT to get the installation id for this repo
    const appOctokit = new Octokit({ auth: jwt });
    const installResp = await appOctokit.request('GET /repos/{owner}/{repo}/installation', {
      owner,
      repo,
    });
    const installationId = installResp.data.id;
    if (!installationId) {
      throw new Error('Could not find installation id for repository');
    }

    // 3) Create an installation access token
    const tokenResp = await appOctokit.request('POST /app/installations/{installation_id}/access_tokens', {
      installation_id: installationId,
    });
    const installationToken = tokenResp.data.token;
    if (!installationToken) throw new Error('Failed to obtain installation access token');

    // 4) Optionally check mergeability / status checks here (recommended)
    const octokit = new Octokit({ auth: installationToken });

    // OPTIONAL: check that PR is mergeable and not draft
    const pr = await octokit.pulls.get({ owner, repo, pull_number });
    if (pr.data.draft) {
      throw new Error('PR is a draft. Aborting merge.');
    }
    // pr.data.mergeable can be null temporarily; rely on branch protection to block merges if checks fail.

    // 5) Merge the PR with squash
    const mergeMethod = 'squash';
    console.log(`Attempting to squash-merge PR #${pull_number} in ${owner}/${repo}...`);
    const mergeResp = await octokit.pulls.merge({
      owner,
      repo,
      pull_number,
      merge_method: mergeMethod,
    });

    console.log('Merge response:', mergeResp.data);
    if (mergeResp.data.merged) {
      console.log(`PR #${pull_number} was squash-merged successfully.`);
      process.exit(0);
    } else {
      console.error('Merge did not complete:', mergeResp.data);
      process.exit(2);
    }
  } catch (err) {
    console.error('Error:', err && err.message ? err.message : err);
    process.exit(1);
  }
}

main();
