// ============================================================
// sync.js — Google Sheets Sync Routes
// Manual trigger + sync status endpoints
// ============================================================

const express = require('express');
const router = express.Router();
const { SyncLog } = require('../db');

// -------------------------------------------------------
// POST /api/sync/trigger — Manually trigger a sync now
// -------------------------------------------------------
router.post('/trigger', async (req, res) => {
  try {
    // Get the syncService that was attached to the app
    const syncService = req.app.get('syncService');
    if (!syncService) {
      return res.status(500).json({ error: 'Sync service not available' });
    }

    const result = await syncService();
    res.json({ success: true, ...result });
  } catch (err) {
    console.error('Manual sync error:', err);
    res.status(500).json({ error: err.message || 'Sync failed' });
  }
});

// -------------------------------------------------------
// GET /api/sync/status — Get the latest sync status
// -------------------------------------------------------
router.get('/status', async (req, res) => {
  try {
    const lastSync = await SyncLog.findOne({ order: [['synced_at', 'DESC']] });
    const history = await SyncLog.findAll({ order: [['synced_at', 'DESC']], limit: 10 });

    res.json({
      last_sync: lastSync,
      history,
      is_configured: !!process.env.GOOGLE_SHEET_ID && process.env.GOOGLE_SHEET_ID !== 'your_google_sheet_id_here',
    });
  } catch (err) {
    res.status(500).json({ error: 'Failed to fetch sync status' });
  }
});

module.exports = router;
