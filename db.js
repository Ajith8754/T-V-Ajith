// ============================================================
// db.js — SQLite Database Setup using Sequelize ORM
// Sequelize with sqlite3 works without Visual Studio build tools
// ============================================================

const { Sequelize, DataTypes } = require('sequelize');
const path = require('path');

// Create database connection using SQLite
const sequelize = new Sequelize({
  dialect: 'sqlite',
  storage: path.join(__dirname, 'tv_data.db'),
  logging: false, // Set to console.log to see all SQL queries
});

// -------------------------------------------------------
// TestReport Model — matches your Excel columns exactly
// -------------------------------------------------------
const TestReport = sequelize.define('TestReport', {
  id: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  sl_no: { type: DataTypes.INTEGER },
  testing_group: { type: DataTypes.STRING },
  test_component: { type: DataTypes.STRING },
  vehicle_model: { type: DataTypes.STRING },
  vin_number: { type: DataTypes.STRING },
  report_number: { type: DataTypes.STRING, unique: true },
  test_name: { type: DataTypes.STRING },
  test_description: { type: DataTypes.TEXT },
  requested_by: { type: DataTypes.STRING },
  test_location: { type: DataTypes.STRING },
  start_date: { type: DataTypes.STRING },
  end_date: { type: DataTypes.STRING },
  report_date: { type: DataTypes.STRING },
  test_engineer: { type: DataTypes.STRING },
  test_decision: { type: DataTypes.STRING },
  test_data: { type: DataTypes.STRING },
  remark: { type: DataTypes.TEXT },
  category: { type: DataTypes.STRING },
  ord_report_number: { type: DataTypes.STRING },
  engineers: { type: DataTypes.STRING },
  raw_data: { type: DataTypes.TEXT },
  source: { type: DataTypes.STRING, defaultValue: 'manual' },
}, {
  tableName: 'test_reports',
  timestamps: true,
  createdAt: 'created_at',
  updatedAt: 'updated_at',
  indexes: [
    { fields: ['testing_group'] },
    { fields: ['vehicle_model'] },
    { fields: ['test_decision'] },
    { fields: ['report_date'] },
    { fields: ['category'] },
  ],
});

// -------------------------------------------------------
// User Model — for login authentication
// -------------------------------------------------------
const User = sequelize.define('User', {
  id: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  username: { type: DataTypes.STRING, unique: true, allowNull: false },
  password_hash: { type: DataTypes.STRING, allowNull: false },
  role: { type: DataTypes.STRING, defaultValue: 'viewer' },
  section: { type: DataTypes.STRING, defaultValue: 'DURABILITY' },
}, {
  tableName: 'users',
  timestamps: true,
  createdAt: 'created_at',
  updatedAt: false,
});

// -------------------------------------------------------
// SyncLog Model — tracks Google Sheets sync history
// -------------------------------------------------------
const SyncLog = sequelize.define('SyncLog', {
  id: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  rows_added: { type: DataTypes.INTEGER, defaultValue: 0 },
  rows_updated: { type: DataTypes.INTEGER, defaultValue: 0 },
  status: { type: DataTypes.STRING, defaultValue: 'success' },
  message: { type: DataTypes.TEXT },
}, {
  tableName: 'sync_log',
  timestamps: true,
  createdAt: 'synced_at',
  updatedAt: false,
});

// -------------------------------------------------------
// Initialize: Create tables if they don't exist
// -------------------------------------------------------
async function initDB() {
  try {
    await sequelize.authenticate();
    await sequelize.sync(); // Creates tables if they don't exist
    
    // Manually add raw_data column if not present to avoid SQLite alter table bug
    try {
      await sequelize.query("ALTER TABLE test_reports ADD COLUMN raw_data TEXT;");
      console.log("✅ Column raw_data successfully added to test_reports");
    } catch (e) {
      // Ignore if duplicate column name error
    }

    console.log('✅ Database connected and tables created: tv_data.db');
  } catch (err) {
    console.error('❌ Database initialization failed:', err.message);
    throw err;
  }
}

module.exports = { sequelize, TestReport, User, SyncLog, initDB };
