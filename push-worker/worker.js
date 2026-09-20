const fs = require('fs');
const path = require('path');
const webpush = require('web-push');

const args = process.argv.slice(2);

const getArg = (name, fallback = '') => {
  const index = args.indexOf(name);
  return index >= 0 && index + 1 < args.length
    ? args[index + 1]
    : fallback;
};

const API_BASE =
  getArg(
    '--api',
    process.env.BUSYCLOUD_PUSH_API ||
      'http://127.0.0.1:8081',
  );

const PARENT_PID = Number(
  getArg('--parent-pid', process.env.BUSYCLOUD_PARENT_PID || '0'),
);

const POLL_MS = Math.max(
  5000,
  Number(process.env.BUSYCLOUD_PUSH_POLL_MS || 10000),
);

const CONFIG_PATH = path.resolve(
  __dirname,
  '..',
  'data',
  'push',
  'push_config.json',
);

const LOCK_PATH = path.resolve(
  __dirname,
  '..',
  'data',
  'push',
  'worker.pid',
);

const sleep = (ms) =>
  new Promise((resolve) => setTimeout(resolve, ms));

const readConfig = () => {
  if (!fs.existsSync(CONFIG_PATH)) {
    throw new Error(
      `Push config not found: ${CONFIG_PATH}. Run setup_push.ps1 first.`,
    );
  }

  const config = JSON.parse(
    fs.readFileSync(CONFIG_PATH, 'utf8'),
  );

  for (const key of [
    'publicKey',
    'privateKey',
    'subject',
    'workerSecret',
  ]) {
    if (!config[key]) {
      throw new Error(
        `Push config is missing required field: ${key}`,
      );
    }
  }

  return config;
};

const acquireSingleInstanceLock = () => {
  fs.mkdirSync(path.dirname(LOCK_PATH), {
    recursive: true,
  });

  if (fs.existsSync(LOCK_PATH)) {
    const oldPid = Number(
      fs.readFileSync(LOCK_PATH, 'utf8').trim(),
    );

    if (oldPid > 0) {
      try {
        process.kill(oldPid, 0);
        console.error(
          `BUSY Cloud push worker already appears to be running (PID ${oldPid}).`,
        );
        process.exit(2);
      } catch {
        // Stale PID file.
      }
    }

    try {
      fs.unlinkSync(LOCK_PATH);
    } catch {}
  }

  fs.writeFileSync(
    LOCK_PATH,
    String(process.pid),
    'utf8',
  );
};

const releaseLock = () => {
  try {
    if (
      fs.existsSync(LOCK_PATH) &&
      fs.readFileSync(LOCK_PATH, 'utf8').trim() ===
        String(process.pid)
    ) {
      fs.unlinkSync(LOCK_PATH);
    }
  } catch {}
};

const minimalPushContent = (job) => {
  const type = String(
    job.notificationType || '',
  ).toUpperCase();

  if (type === 'WEB_APPROVAL_SUBMITTED') {
    return {
      title: 'New Web Approval',
      body: 'A transaction is waiting for your approval.',
      url: job.webApprovalId
        ? `/web-approvals?id=${encodeURIComponent(
            job.webApprovalId,
          )}`
        : '/web-approvals',
    };
  }

  if (type === 'WEB_APPROVAL_APPROVED') {
    return {
      title: 'Voucher Approved',
      body: 'Your Web Approval transaction was approved.',
      url: '/',
    };
  }

  if (type === 'WEB_APPROVAL_REJECTED') {
    return {
      title: 'Voucher Rejected',
      body: 'Your Web Approval transaction was rejected. Sign in to view details.',
      url: '/',
    };
  }

  if (type === 'WEB_APPROVAL_SYNCED') {
    return {
      title: 'Voucher Posted to BUSY',
      body: 'Your approved transaction was synchronized to BUSY.',
      url: '/',
    };
  }

  if (
    type === 'WEB_APPROVAL_SYNC_REVIEW_REQUIRED'
  ) {
    return {
      title: 'Web Approval Needs Review',
      body: 'A BUSY synchronization requires review. Automatic retry is blocked.',
      url: job.webApprovalId
        ? `/web-approvals?id=${encodeURIComponent(
            job.webApprovalId,
          )}`
        : '/web-approvals',
    };
  }

  if (type === 'WEB_APPROVAL_SYNC_FAILED') {
    return {
      title: 'Web Approval Sync Failed',
      body: 'A synchronization attempt failed. Sign in to review the transaction.',
      url: job.webApprovalId
        ? `/web-approvals?id=${encodeURIComponent(
            job.webApprovalId,
          )}`
        : '/web-approvals',
    };
  }

  return {
    title: 'BUSY Cloud',
    body: 'You have a new notification.',
    url: '/',
  };
};

const requestJson = async (
  url,
  options = {},
  workerSecret,
) => {
  const response = await fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      'Content-Type': 'application/json',
      'X-BusyCloud-Push-Worker': workerSecret,
    },
  });

  const text = await response.text();
  let data = null;

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = {
      success: false,
      error: text || `HTTP ${response.status}`,
    };
  }

  if (!response.ok || data?.success === false) {
    throw new Error(
      data?.error ||
        `Push worker API request failed (${response.status}).`,
    );
  }

  return data;
};

const reportResult = async (
  config,
  job,
  result,
) => {
  try {
    await requestJson(
      `${API_BASE}/busy/push-worker/result`,
      {
        method: 'POST',
        body: JSON.stringify({
          instanceId: job.instanceId,
          companyCode: job.companyCode,
          deliveryId: job.deliveryId,
          subscriptionId: job.subscriptionId,
          succeeded: Boolean(result.succeeded),
          gone: Boolean(result.gone),
          httpStatus: Number(result.httpStatus || 0),
          error: result.error || '',
        }),
      },
      config.workerSecret,
    );
  } catch (error) {
    console.error(
      `[PUSH] Could not report delivery ${job.deliveryId}:`,
      error.message,
    );
  }
};

const sendJob = async (config, job) => {
  const content = minimalPushContent(job);

  const payload = JSON.stringify({
    title: content.title,
    body: content.body,
    requireInteraction:
      job.notificationType ===
      'WEB_APPROVAL_SYNC_REVIEW_REQUIRED',
    data: {
      notificationId: job.notificationId,
      webApprovalId: job.webApprovalId || '',
      type: job.notificationType || '',
      url: content.url,
    },
  });

  try {
    const response = await webpush.sendNotification(
      job.subscription,
      payload,
      {
        TTL: 60 * 60 * 24,
        urgency: 'normal',
      },
    );

    console.log(
      `[PUSH SENT] ${job.instanceId}/${job.companyCode} ` +
        `${job.notificationType} -> ${response.statusCode}`,
    );

    await reportResult(config, job, {
      succeeded: true,
      gone: false,
      httpStatus: response.statusCode,
      error: '',
    });
  } catch (error) {
    const statusCode = Number(
      error?.statusCode || error?.status || 0,
    );

    const gone =
      statusCode === 404 || statusCode === 410;

    console.warn(
      `[PUSH ${gone ? 'GONE' : 'FAILED'}] ` +
        `${job.instanceId}/${job.companyCode} ` +
        `${job.notificationType} -> ${statusCode || 'no status'} ` +
        `${error?.message || error}`,
    );

    await reportResult(config, job, {
      succeeded: false,
      gone,
      httpStatus: statusCode,
      error: String(
        error?.body ||
          error?.message ||
          error ||
          'Web Push delivery failed.',
      ).slice(0, 4000),
    });
  }
};

let polling = false;
let shuttingDown = false;
let lastBackendWarningAt = 0;

const parentIsAlive = () => {
  if (!PARENT_PID || PARENT_PID <= 0) {
    return true;
  }

  try {
    process.kill(PARENT_PID, 0);
    return true;
  } catch {
    return false;
  }
};

const poll = async (config) => {
  if (polling || shuttingDown) return;
  polling = true;

  try {
    const data = await requestJson(
      `${API_BASE}/busy/push-worker/jobs?limit=50`,
      { method: 'GET' },
      config.workerSecret,
    );

    const jobs = Array.isArray(data?.data?.jobs)
      ? data.data.jobs
      : [];

    for (const job of jobs) {
      if (shuttingDown) break;
      await sendJob(config, job);
    }
  } catch (error) {
    const now = Date.now();

    // Avoid filling logs while the PowerShell bridge is restarting.
    if (now - lastBackendWarningAt > 30000) {
      console.warn(
        `[PUSH WORKER] Backend unavailable or rejected worker request: ${error.message}`,
      );
      lastBackendWarningAt = now;
    }
  } finally {
    polling = false;
  }
};

const main = async () => {
  acquireSingleInstanceLock();

  const config = readConfig();

  webpush.setVapidDetails(
    config.subject,
    config.publicKey,
    config.privateKey,
  );

  console.log(
    `[PUSH WORKER] Started. API=${API_BASE} poll=${POLL_MS}ms`,
  );

  while (!shuttingDown) {
    if (!parentIsAlive()) {
      console.log(
        `[PUSH WORKER] Parent process ${PARENT_PID} exited. Stopping worker.`,
      );
      break;
    }

    await poll(config);
    await sleep(POLL_MS);
  }

  releaseLock();
};

const shutdown = () => {
  if (shuttingDown) return;
  shuttingDown = true;
  releaseLock();
  process.exit(0);
};

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
process.on('exit', releaseLock);

main().catch((error) => {
  console.error('[PUSH WORKER FATAL]', error);
  releaseLock();
  process.exit(1);
});
