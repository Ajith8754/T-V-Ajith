// ============================================================
// server.js — Main Express Server
// This is the entry point for the backend
// Run with: node server.js
// ============================================================

require('dotenv').config({ path: require('path').join(__dirname, '.env') }); // Load environment variables from .env file
const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const { initDB, TestReport, SyncLog } = require('./db');
const { readFromGoogleSheet } = require('./services/googleSheets');
const authRoutes = require('./routes/auth');
const reportsRoutes = require('./routes/reports');
const syncRoutes = require('./routes/sync');

const app = express();
const httpServer = http.createServer(app);

// -------------------------------------------------------
// Socket.io — Real-time live updates
// Notifies all browsers when data changes
// -------------------------------------------------------
const io = new Server(httpServer, {
  cors: {
    origin: (origin, callback) => {
      // Allow any localhost port (5173, 5174, 5175, 3000, etc.)
      if (!origin || origin.match(/^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/)) {
        callback(null, true);
      } else {
        callback(new Error('Not allowed by CORS'));
      }
    },
    credentials: true,
  },
});

// Store io instance on app for use in routes
app.set('io', io);

io.on('connection', (socket) => {
  console.log(`🔌 Client connected: ${socket.id}`);
  socket.on('disconnect', () => {
    console.log(`🔌 Client disconnected: ${socket.id}`);
  });
});

// -------------------------------------------------------
// Middleware
// -------------------------------------------------------
app.use(cors({
  origin: (origin, callback) => {
    // Allow any localhost port (5173, 5174, 5175, 3000, etc.)
    if (!origin || origin.match(/^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/)) {
      callback(null, true);
    } else {
      callback(new Error('Not allowed by CORS'));
    }
  },
  credentials: true,
}));
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true }));

// Serve static uploads
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}
app.use('/uploads', express.static(uploadsDir));

// -------------------------------------------------------
// Routes
// -------------------------------------------------------
app.use('/api/auth', authRoutes);
app.use('/api/reports', reportsRoutes);
app.use('/api/sync', syncRoutes);

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// -------------------------------------------------------
// Google Sheets Sync Service
// Runs every 60 seconds (configurable via SYNC_INTERVAL_MS)
// -------------------------------------------------------
async function syncWithGoogleSheets() {
  const isConfigured = process.env.GOOGLE_SHEET_ID && 
                       process.env.GOOGLE_SHEET_ID !== 'your_google_sheet_id_here';

  if (!isConfigured) {
    console.log('⚠️  Google Sheets not configured. Skipping sync. (See SETUP_GUIDE.md)');
    return { status: 'skipped', message: 'Google Sheets not configured' };
  }

  console.log('🔄 Syncing with Google Sheets...');

  try {
    const sheetRows = await readFromGoogleSheet();
    let rowsAdded = 0, rowsUpdated = 0;

    for (const row of sheetRows) {
      if (!row.report_number) continue;

      const existing = await TestReport.findOne({ where: { report_number: row.report_number } });
      if (existing) {
        let newSource = row.source;
        if (existing.source && existing.source !== row.source) {
          const sourcesSet = new Set(existing.source.split(',').map(s => s.trim()));
          sourcesSet.add(row.source);
          newSource = Array.from(sourcesSet).join(',');
        }
        await existing.update({
          ...row,
          source: newSource
        });
        rowsUpdated++;
      } else {
        await TestReport.create(row);
        rowsAdded++;
      }
    }

    // Log sync
    await SyncLog.create({
      rows_added: rowsAdded,
      rows_updated: rowsUpdated,
      status: 'success',
      message: `Synced ${sheetRows.length} rows from Google Sheets`,
    });

    // Notify all connected browsers of the update
    io.emit('data_updated', {
      type: 'google_sheets_sync',
      rows_added: rowsAdded,
      rows_updated: rowsUpdated,
      timestamp: new Date().toISOString(),
    });

    console.log(`✅ Sync complete: ${rowsAdded} added, ${rowsUpdated} updated`);
    return { success: true, rows_added: rowsAdded, rows_updated: rowsUpdated };

  } catch (err) {
    console.error('❌ Sync failed:', err.message);
    await SyncLog.create({
      rows_added: 0,
      rows_updated: 0,
      status: 'error',
      message: err.message,
    });
    throw err;
  }
}

// Attach sync service to app for manual trigger
app.set('syncService', syncWithGoogleSheets);

// -------------------------------------------------------
// Database initialization + Server start
// -------------------------------------------------------
const PORT = process.env.PORT || 5000;

async function startServer() {
  // 1. Initialize database and create tables
  await initDB();

  // 2. Create default admin user if not exists
  const bcrypt = require('bcryptjs');
  const { User } = require('./db');

  // Clean up old default user if exists
  await User.destroy({ where: { username: 'Ajith123' } });

  const existing = await User.findOne({ where: { username: 'simple123' } });
  if (!existing) {
    const hash = await bcrypt.hash('123admin', 10);
    await User.create({ username: 'simple123', password_hash: hash, role: 'admin', section: 'ALL' });
    console.log('✅ Default user created: simple123');
  }

  // Clean up any stale preview/upload data in database on startup
  const { Op } = require('sequelize');
  await TestReport.destroy({
    where: {
      [Op.or]: [
        { source: 'google_sheets:upload data' },
        { source: { [Op.like]: 'upload%' } }
      ]
    }
  });

  // 3. Start Google Sheets sync on interval (disabled by default to prevent overwriting local uploads, use manual sync in navbar)
  const syncIntervalMs = parseInt(process.env.SYNC_INTERVAL_MS) || 60000;

  // Run initial sync after 3 seconds (give server time to fully start)
  setTimeout(() => {
    syncWithGoogleSheets().catch(() => {});
  }, 3000);

  // Then repeat on schedule
  setInterval(() => {
    syncWithGoogleSheets().catch(() => {});
  }, syncIntervalMs);

  // 4. Start HTTP server
  httpServer.listen(PORT, () => {
    console.log('');
    console.log('╔════════════════════════════════════════════╗');
    console.log('║  T&V Simple Energy Backend Server          ║');
    console.log(`║  Running on: http://localhost:${PORT}          ║`);
    console.log(`║  Google Sync: every ${syncIntervalMs / 1000}s                    ║`);
    console.log('╚════════════════════════════════════════════╝');
    console.log('');
  });
}

startServer().catch(err => {
  console.error('Failed to start server:', err);
  process.exit(1);
});

// -------------------------------------------------------
// Global Error Handlers — keep server alive through minor errors
// -------------------------------------------------------

// Ignore harmless client-disconnect errors (browser closed tab, redirected, etc.)
const IGNORABLE_ERRORS = new Set(['ECONNABORTED', 'ECONNRESET', 'EPIPE', 'ETIMEDOUT']);

process.on('uncaughtException', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`\n❌ Port ${PORT} is already in use!`);
    console.error('   Run this command to free it:');
    console.error(`   Get-NetTCPConnection -LocalPort ${PORT} | ForEach-Object { taskkill /PID $_.OwningProcess /F }\n`);
    process.exit(1);
  } else if (IGNORABLE_ERRORS.has(err.code)) {
    // Client disconnected mid-request — not a real error, ignore it
    return;
  } else {
    console.error('❌ Unexpected server error:', err.message);
    // Don't exit — keep server running
  }
});

process.on('unhandledRejection', (reason) => {
  // Log but don't crash the server
  console.error('⚠️  Unhandled promise rejection:', reason?.message || reason);
});

