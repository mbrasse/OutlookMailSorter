// scripts/merge-pr-checks.js
// Safer merge script for GitHub App (CommonJS friendly, uses dynamic ESM imports):
// - Uses dynamic import() to consume modern @octokit packages without require() errors.
// - Requires env: APP_ID, PRIVATE_KEY, GITHUB_REPOSITORY, PULL_NUMBER
// - Checks PR draft state, required approvals, and (optionally) status checks before merging.
// - Default merge method: squash
//
// Commit message update: Switch to ESM dynamic import usage to avoid require() errors with modern @octokit packages.
// Retains previous behavior and CLI interface.

'use strict';

async function dynamicImportOctokit() {
  // Dynamically import both packages; this avoids require() issues in environments where
  // packages are published as ESM-only.
  const appMod = await import('@octokit/app');
  const restMod = await import('@octokit/rest');
  // Named exports expected
  const App = appMod.App || appMod.default;
  const Octokit = restMod.Octokit || restMod.default;
  if (!App) throw new Error('Failed to import App from @octokit/app');
  if (!Octokit) throw new Error('Failed to import Octokit from @octokit/rest');
  return { App, Octokit };
}

async function getInstallationToken(appId, privateKey, owner, repo) {
  const { App, Octokit } = await dynamicImportOctokit();
  const app = new App({ appId: Number(appId), privateKey });
  const jwt = app.getSignedJsonWebToken();
  const appOctokit = new Octokit({ auth: jwt });

  // Get installation for this repo
  const installResp = await appOctokit.request('GET /repos/{owner}/{repo}/installation', { owner, repo });
  const installationId = installResp && installResp.data && installResp.data.id;
  if (!installationId) throw new Error('Could not find installation id for repository');
  const tokenResp = await appOctokit.request('POST /app/installations/{installation_id}/access_tokens', {
    installation_id: installationId,
  });
  if (!tokenResp || !tokenResp.data || !tokenResp.data.token) throw new Error('Failed to obtain installation access token');
  return tokenResp.data.token;
}

async function waitForStatusChecks(octokit, owner, repo, ref, requiredContexts = [], timeoutMs = 5 * 60 * 1000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    // Use combined status API
    const resp = await octokit.repos.getCombinedStatusForRef({ owner, repo, ref });
    const state = resp && resp.data && resp.data.state; // 'success', 'pending', 'failure'
    if (requiredContexts && requiredContexts.length > 0) {
      const contexts = (resp && resp.data && resp.data.statuses) || [];
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
  const reviews = (reviewsResp && reviewsResp.data) || [];
  // Only consider the latest review by each user
  const latestByUser = new Map();
  for (const r of reviews) {
    if (r && r.user && r.user.login) {
      latestByUser.set(r.user.login, r);
    }
  }
  const approved = [...latestByUser.values()].filter(r => r.state === 'APPROVED').length;
  return approved >= requiredApprovals;
}

async function mergeIfReady({ appId, privateKey, owner, repo, pull_number, mergeMethod = 'squash', requiredContexts = [], requiredApprovals = 1 }) {
  if (!appId) throw new Error('APP_ID missing');
  if (!privateKey) throw new Error('PRIVATE_KEY missing');
  if (!owner || !repo) throw new Error('Repository owner or name missing');

  // Normalize private key newlines if stored with \n
  privateKey = privateKey.replace(/\\n/g, '\n');

  const token = await getInstallationToken(appId, privateKey, owner, repo);

  const { Octokit } = await dynamicImportOctokit();
  const octokit = new Octokit({ auth: token });

  const prResp = await octokit.pulls.get({ owner, repo, pull_number });
  const pr = prResp && prResp.data;
  if (!pr) throw new Error('PR not found');

  if (pr.draft) throw new Error('PR is a draft');

  // Wait for mergeable state to be computed (GitHub may return null initially)
  let attempts = 0;
  let currentPr = pr;
  while ((currentPr.mergeable === null || currentPr.mergeable === undefined) && attempts < 10) {
    await new Promise(r => setTimeout(r, 2000));
    const fresh = await octokit.pulls.get({ owner, repo, pull_number });
    if (fresh && fresh.data) {
      currentPr = fresh.data;
      if (currentPr.mergeable !== null && currentPr.mergeable !== undefined) break;
    }
    attempts++;
  }
  if (currentPr.mergeable === false) throw new Error(`PR is not mergeable: ${currentPr.mergeable_state || 'unknown'}`);

  // Optionally wait for required status checks (if any)
  if (requiredContexts && requiredContexts.length > 0) {
    await waitForStatusChecks(octokit, owner, repo, currentPr.head.sha, requiredContexts);
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
  return mergeResp && mergeResp.data;
}

// CLI entrypoint (CommonJS compatible)
if (require && require.main === module) {
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
      if (result && result.merged) process.exit(0);
      else process.exit(2);
    } catch (err) {
      console.error('Error:', err && err.message ? err.message : err);
      process.exit(1);
    }
  })();
}
