#!/bin/bash

###############################################################################
# HR PAYROLL DEPLOYMENT SCRIPT
# Automated deployment to Oracle Cloud
###############################################################################

set -e  # Exit on any error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Configuration
APP_DIR="/var/www/hr-payroll"
APP_NAME="hr-payroll-api"
BACKUP_DIR="/var/backups/hr-payroll"

echo ""
echo "════════════════════════════════════════════════════"
echo "   HR PAYROLL API - DEPLOYMENT"
echo "════════════════════════════════════════════════════"
echo ""

# Check if running as correct user
if [ "$USER" != "ubuntu" ] && [ "$USER" != "root" ]; then
  echo -e "${RED}❌ Please run this script as ubuntu user${NC}"
  exit 1
fi

# Navigate to application directory
echo "📂 Navigating to $APP_DIR..."
cd "$APP_DIR" || exit 1

# Backup current version (if exists)
if [ -d "$APP_DIR/node_modules" ]; then
  echo ""
  echo "💾 Creating backup of current version..."
  TIMESTAMP=$(date +%Y%m%d_%H%M%S)
  BACKUP_FILE="$BACKUP_DIR/app_backup_$TIMESTAMP.tar.gz"
  mkdir -p "$BACKUP_DIR"

  tar -czf "$BACKUP_FILE" \
    --exclude='node_modules' \
    --exclude='.git' \
    --exclude='*.log' \
    .

  echo -e "${GREEN}✅ Backup created: $BACKUP_FILE${NC}"
fi

# Pull latest changes (if using git)
if [ -d ".git" ]; then
  echo ""
  echo "📥 Pulling latest changes from git..."
  git pull origin main || git pull origin master
  echo -e "${GREEN}✅ Git pull complete${NC}"
fi

# Install/update dependencies
echo ""
echo "📦 Installing/updating dependencies..."
npm install --production
echo -e "${GREEN}✅ Dependencies installed${NC}"

# Run tests (if available)
if [ -f "package.json" ] && grep -q "\"test\"" package.json; then
  echo ""
  echo "🧪 Running tests..."
  npm test || echo -e "${YELLOW}⚠️  Tests failed or not configured${NC}"
fi

# Restart application with PM2
echo ""
echo "🔄 Restarting application..."

# Check if PM2 process exists
if pm2 describe "$APP_NAME" > /dev/null 2>&1; then
  echo "   Restarting existing PM2 process..."
  pm2 restart "$APP_NAME" --update-env
else
  echo "   Starting new PM2 process..."
  pm2 start server.js --name "$APP_NAME"
fi

# Save PM2 configuration
pm2 save

echo -e "${GREEN}✅ Application restarted${NC}"

# Show status
echo ""
echo "📊 Application status:"
pm2 describe "$APP_NAME"

echo ""
echo "📋 Recent logs:"
pm2 logs "$APP_NAME" --lines 10 --nostream

echo ""
echo "════════════════════════════════════════════════════"
echo -e "${GREEN}✅ DEPLOYMENT COMPLETED SUCCESSFULLY!${NC}"
echo "════════════════════════════════════════════════════"
echo ""
echo "Next steps:"
echo "  • Test API: curl http://localhost:3000/health"
echo "  • View logs: pm2 logs $APP_NAME"
echo "  • Monitor: pm2 monit"
echo ""

exit 0
