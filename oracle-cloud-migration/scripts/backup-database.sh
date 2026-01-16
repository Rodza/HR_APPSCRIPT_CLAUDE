#!/bin/bash

###############################################################################
# HR PAYROLL DATABASE BACKUP SCRIPT
# Automated MySQL backup with compression and retention
###############################################################################

# Configuration
DB_USER="hrapp"
DB_NAME="hrdb"
DB_PASSWORD="${DB_PASSWORD:-}"  # Set via environment variable for security

BACKUP_DIR="/var/backups/hr-payroll"
RETENTION_DAYS=30  # Keep backups for 30 days

# Create backup directory if it doesn't exist
mkdir -p "$BACKUP_DIR"

# Generate backup filename with timestamp
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
BACKUP_FILE="$BACKUP_DIR/hrdb_backup_$TIMESTAMP.sql"
COMPRESSED_FILE="$BACKUP_FILE.gz"

echo "════════════════════════════════════════════════════"
echo "  HR Payroll Database Backup"
echo "════════════════════════════════════════════════════"
echo "Started: $(date)"
echo "Database: $DB_NAME"
echo ""

# Check if password is provided
if [ -z "$DB_PASSWORD" ]; then
  echo "⚠️  DB_PASSWORD not set as environment variable"
  echo "   Reading password interactively..."
  echo ""
fi

# Perform backup
echo "📦 Creating backup..."
if [ -z "$DB_PASSWORD" ]; then
  # Interactive password
  mysqldump -u "$DB_USER" -p "$DB_NAME" > "$BACKUP_FILE"
else
  # Password from environment
  mysqldump -u "$DB_USER" -p"$DB_PASSWORD" "$DB_NAME" > "$BACKUP_FILE"
fi

# Check if backup was successful
if [ $? -eq 0 ]; then
  echo "✅ Backup created: $BACKUP_FILE"

  # Get backup size
  BACKUP_SIZE=$(du -h "$BACKUP_FILE" | cut -f1)
  echo "   Size: $BACKUP_SIZE"

  # Compress backup
  echo ""
  echo "🗜️  Compressing backup..."
  gzip "$BACKUP_FILE"

  COMPRESSED_SIZE=$(du -h "$COMPRESSED_FILE" | cut -f1)
  echo "✅ Compressed: $COMPRESSED_FILE"
  echo "   Size: $COMPRESSED_SIZE"

  # Remove old backups
  echo ""
  echo "🧹 Removing backups older than $RETENTION_DAYS days..."
  find "$BACKUP_DIR" -name "hrdb_backup_*.sql.gz" -mtime +$RETENTION_DAYS -delete

  # List remaining backups
  BACKUP_COUNT=$(find "$BACKUP_DIR" -name "hrdb_backup_*.sql.gz" | wc -l)
  echo "   $BACKUP_COUNT backup(s) retained"

  echo ""
  echo "════════════════════════════════════════════════════"
  echo "✅ Backup completed successfully!"
  echo "════════════════════════════════════════════════════"
  echo "Finished: $(date)"

else
  echo ""
  echo "════════════════════════════════════════════════════"
  echo "❌ Backup failed!"
  echo "════════════════════════════════════════════════════"
  exit 1
fi

# Exit successfully
exit 0
