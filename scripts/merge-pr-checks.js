// scripts/merge-pr-checks.js
// Safer merge script for GitHub App:
// - Requires env: APP_ID, PRIVATE_KEY, GITHUB_REPOSITORY, PULL_NUMBER
// - Checks PR draft state, required approvals, and (optionally) status checks before merging.

const { App } = require('@octokit/app');
const { Octokit } = require('@octokit/rest');

async function getInstallationToken(appId, privateKey, owner, repo) {
  const app = new App({ id: Number(appId), privateKey });
  const jwt = app.getSignedJsonWebToken();
  const appOctokit = new Octokit({ auth: jwt });
  // Get installation for this repo
  const installResp = await appOctokit.request('GET /repos/{owner}/{repo}/installation', { owner, repo });
  const installationId = installResp.data.id;
  if (!installationId) throw new Error('Could not find installation id for repository');
  const tokenResp = await appOctokit.request('POST /app/installations/{installation_id}/access_tokens', {
    installation_id: installationId,
  });
  if (!tokenResp.data || !tokenResp.data.token) throw new Error('Failed to obtain installation access token');
  return tokenResp.data.token;
}

async function waitForStatusChecks(octokit, owner, repo, ref, requiredContexts = [], timeoutMs = 5 * 60 * 1000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    // Use combined status API
    const resp = await octokit.repos.getCombinedStatusForRef({ owner, repo, ref });
    const state = resp.data.state; // 'success', 'pending', 'failure'
    if (requiredContexts.length > 0) {
      const contexts = resp.data.statuses || [];
      const ok = requiredContexts.every(ctx => contexts.find(s => s.context === ctx && s.state === 'success'));
      if (ok) return true;
    } else {
      if (state === 'success') return true;
    }
    // wait a bit
    await new Promise(r => setTimeout(r, 5000));
  }
  throw new Error('Timeout waiting for status checks to pass');
}

async function ensureApprovals(octokit, owner, repo, pull_number, requiredApprovals = 1) {
  const reviewsResp = await octokit.pulls.listReviews({ owner, repo, pull_number });
  // Only consider the latest review by each user
  const latestByUser = new Map();
  for (const r of reviewsResp.data) {
    latestByUser.set(r.user.login, r);
  }
  const approved = [...latestByUser.values()].filter(r => r.state === 'APPROVED').length;
  return approved >= requiredApprovals;
}

async function mergeIfReady({ appId, privateKey, owner, repo, pull_number, mergeMethod = 'squash', requiredContexts = [], requiredApprovals = 1 }) {
  // Normalize private key newlines if stored with \n
  privateKey = privateKey.replace(/\\n/g, '\n');

  const token = await getInstallationToken(appId, privateKey, owner, repo);
  const octokit = new Octokit({ auth: token });

  const prResp = await octokit.pulls.get({ owner, repo, pull_number });
  const pr = prResp.data;
  if (!pr) throw new Error('PR not found');

  if (pr.draft) throw new Error('PR is a draft');

  // Wait for mergeable state to be computed (GitHub may return null initially)
  let attempts = 0;
  while (pr.mergeable === null && attempts < 10) {
    await new Promise(r => setTimeout(r, 2000));
    const fresh = await octokit.pulls.get({ owner, repo, pull_number });
    if (fresh.data.mergeable !== null) {
      pr.mergeable = fresh.data.mergeable;
      pr.mergeable_state = fresh.data.mergeable_state;
      break;
    }
    attempts++;
  }
  if (pr.mergeable === false) throw new Error(`PR is not mergeable: ${pr.mergeable_state}`);

  // Optionally wait for required status checks (if any)
  if (requiredContexts && requiredContexts.length > 0) {
    await waitForStatusChecks(octokit, owner, repo, pr.head.sha, requiredContexts);
  }

  // Ensure approvals
  const approvalsOk = await ensureApprovals(octokit, owner, repo, pull_number, requiredApprovals);
  if (!approvalsOk) throw new Error('PR does not have required approvals');

  // Finally attempt the merge
  const mergeResp = await octokit.pulls.merge({
    owner,
    repo,
    pull_number,
    merge_method: mergeMethod,
  });
  return mergeResp.data;
}

// CLI entrypoint
if (require.main === module) {
  (async () => {
    try {
      const appId = process.env.APP_ID;
      let privateKey = process.env.PRIVATE_KEY;
      const repository = process.env.GITHUB_REPOSITORY;
      const pullNumberEnv = process.env.PULL_NUMBER || '';

      if (!appId) throw new Error('APP_ID missing');
      if (!privateKey) throw new Error('PRIVATE_KEY missing');
      if (!repository) throw new Error('GITHUB_REPOSITORY missing');

      const [owner, repo] = repository.split('/');
      const pull_number = Number(pullNumberEnv);
      if (!pull_number || Number.isNaN(pull_number)) throw new Error(`Invalid PULL_NUMBER: "${pullNumberEnv}"`);

      // Customize these as needed: requiredContexts = [] -> rely on branch protection
      const requiredContexts = []; // e.g. ['ci/build', 'ci/test']
      const requiredApprovals = 1; // rely on CODEOWNERS and branch protection

      const result = await mergeIfReady({
        appId,
        privateKey,
        owner,
        repo,
        pull_number,
        mergeMethod: 'squash',
        requiredContexts,
        requiredApprovals,
      });

      console.log('Merge result:', result);
      if (result.merged) process.exit(0);
      else process.exit(2);
    } catch (err) {
      console.error('Error:', err && err.message ? err.message : err);
      process.exit(1);
    }
  })();
}
