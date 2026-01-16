# Oracle Cloud Always Free - Complete Setup Guide
## HR Payroll System Migration

This guide walks you through setting up the HR Payroll system on Oracle Cloud's **Always Free** tier ($0 forever!).

---

## 📋 Table of Contents

1. [Create Oracle Cloud Account](#step-1-create-oracle-cloud-account)
2. [Create ARM VM Instance](#step-2-create-arm-vm-instance)
3. [Configure Network & Firewall](#step-3-configure-network--firewall)
4. [Install MySQL](#step-4-install-mysql)
5. [Install Node.js](#step-5-install-nodejs)
6. [Setup Database](#step-6-setup-database)
7. [Deploy Application](#step-7-deploy-application)
8. [Configure Auto-Start](#step-8-configure-auto-start)
9. [Setup SSL (Optional)](#step-9-setup-ssl-optional)
10. [Testing](#step-10-testing)

---

## Step 1: Create Oracle Cloud Account

### 1.1 Sign Up

1. Go to: https://signup.oraclecloud.com/
2. Click "Start for free"
3. Fill in details:
   - Email address
   - Country (select your country)
   - Name and company info
4. **Credit card required** (for verification only - you won't be charged)
5. Choose **home region carefully** (e.g., US East, South Africa)
   - ⚠️ You cannot change this later!
   - Always Free resources only available in home region

### 1.2 Verify Account

1. Check email for verification link
2. Complete email verification
3. Wait for account activation (usually instant, can take up to 24 hours)

### 1.3 Sign In

1. Go to: https://cloud.oracle.com/
2. Enter cloud account name (from email)
3. Sign in with your credentials

---

## Step 2: Create ARM VM Instance

### 2.1 Navigate to Compute

1. Click hamburger menu (☰) top left
2. Go to: **Compute** → **Instances**
3. Click "Create Instance"

### 2.2 Configure Instance

**Name:**
```
hr-payroll-server
```

**Placement:**
- Availability Domain: (keep default)
- Fault Domain: (keep default)

**Image and Shape:**

1. Click "Edit" next to Image and shape
2. Click "Change Shape"
3. Select **"Ampere"** (ARM processor)
4. Select **"VM.Standard.A1.Flex"**
5. Configure resources:
   - **OCPUs:** 2 (out of your 4 free)
   - **Memory:** 12 GB (out of your 24 GB free)
   - This leaves room for future services!

**Image:**
- Click "Change Image"
- Select **"Canonical Ubuntu 22.04"** or latest LTS
- Check "I have reviewed and accept the terms"

**Networking:**

- VCN: (create new if none exists)
  - Name: `hr-payroll-vcn`
  - Subnet: Public subnet (default)
- **Assign a public IPv4 address:** ✅ CHECK THIS!

**Add SSH Keys:**

1. **Option A - Generate new key pair:** (Recommended for beginners)
   - Select "Generate a key pair for me"
   - Click "Save Private Key" → save to safe location!
   - Click "Save Public Key" → save to safe location!

2. **Option B - Use existing key:**
   - Select "Upload public key files (.pub)"
   - Upload your `id_rsa.pub` file

**Boot volume:**
- Keep defaults (50 GB)

### 2.3 Create Instance

1. Click "Create"
2. Wait for provisioning (2-5 minutes)
3. Status will change from "Provisioning" → "Running"
4. **Copy the Public IP address** (e.g., 132.145.123.45)

---

## Step 3: Configure Network & Firewall

### 3.1 Configure Security List (Oracle Cloud Firewall)

1. On instance details page, click on the **Subnet** link
2. Click on the **Security List** (e.g., "Default Security List")
3. Click "Add Ingress Rules"

**Add these rules:**

**Rule 1 - SSH (Port 22):**
```
Source Type: CIDR
Source CIDR: 0.0.0.0/0
IP Protocol: TCP
Source Port Range: (leave empty)
Destination Port Range: 22
Description: SSH access
```

**Rule 2 - HTTP (Port 80):**
```
Source Type: CIDR
Source CIDR: 0.0.0.0/0
IP Protocol: TCP
Destination Port Range: 80
Description: HTTP access
```

**Rule 3 - HTTPS (Port 443):**
```
Source Type: CIDR
Source CIDR: 0.0.0.0/0
IP Protocol: TCP
Destination Port Range: 443
Description: HTTPS access
```

**Rule 4 - Node.js API (Port 3000):**
```
Source Type: CIDR
Source CIDR: 0.0.0.0/0
IP Protocol: TCP
Destination Port Range: 3000
Description: HR Payroll API
```

### 3.2 Connect to VM via SSH

**On Windows (use PuTTY or Windows Terminal):**
```bash
ssh -i C:\path\to\private-key.key ubuntu@YOUR_PUBLIC_IP
```

**On Mac/Linux:**
```bash
chmod 400 ~/path/to/private-key.key
ssh -i ~/path/to/private-key.key ubuntu@YOUR_PUBLIC_IP
```

First time will ask to accept fingerprint - type `yes`

You should see:
```
ubuntu@hr-payroll-server:~$
```

### 3.3 Configure Ubuntu Firewall (iptables)

Oracle Cloud uses iptables which blocks ports by default.

```bash
# Allow port 3000 (API)
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 3000 -j ACCEPT

# Allow port 80 (HTTP)
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 80 -j ACCEPT

# Allow port 443 (HTTPS)
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 443 -j ACCEPT

# Save rules
sudo netfilter-persistent save
```

If `netfilter-persistent` not found:
```bash
sudo apt install iptables-persistent
sudo netfilter-persistent save
```

---

## Step 4: Install MySQL

### 4.1 Update System

```bash
sudo apt update
sudo apt upgrade -y
```

### 4.2 Install MySQL Server

```bash
sudo apt install mysql-server -y
```

### 4.3 Secure MySQL

```bash
sudo mysql_secure_installation
```

Answer prompts:
```
VALIDATE PASSWORD COMPONENT? Y
Password strength: 2 (strong)
New password: [create strong password - save it!]
Re-enter: [same password]
Remove anonymous users? Y
Disallow root login remotely? Y
Remove test database? Y
Reload privilege tables? Y
```

### 4.4 Verify MySQL Running

```bash
sudo systemctl status mysql
```

Should show: `active (running)`

---

## Step 5: Install Node.js

### 5.1 Install Node.js 18 LTS

```bash
curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
sudo apt install -y nodejs
```

### 5.2 Verify Installation

```bash
node --version  # Should show v18.x.x
npm --version   # Should show 9.x.x or higher
```

### 5.3 Install PM2 (Process Manager)

```bash
sudo npm install -g pm2
```

---

## Step 6: Setup Database

### 6.1 Create Database and User

```bash
sudo mysql
```

In MySQL prompt:
```sql
-- Create database
CREATE DATABASE hrdb CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

-- Create user
CREATE USER 'hrapp'@'localhost' IDENTIFIED BY 'YourSecurePassword123!';

-- Grant privileges
GRANT ALL PRIVILEGES ON hrdb.* TO 'hrapp'@'localhost';

-- Apply changes
FLUSH PRIVILEGES;

-- Verify
SHOW DATABASES;

-- Exit
EXIT;
```

### 6.2 Import Schema

```bash
# Upload schema.sql to server (from your local machine)
# Option A: Use SCP
scp -i ~/private-key.key /path/to/schema.sql ubuntu@YOUR_IP:~/

# Then on server:
mysql -u hrapp -p hrdb < ~/schema.sql
# Enter password when prompted
```

### 6.3 Verify Schema

```bash
mysql -u hrapp -p hrdb
```

```sql
SHOW TABLES;
-- Should show: employees, mastersalary, employee_loans, etc.

EXIT;
```

---

## Step 7: Deploy Application

### 7.1 Create Application Directory

```bash
sudo mkdir -p /var/www/hr-payroll
sudo chown ubuntu:ubuntu /var/www/hr-payroll
cd /var/www/hr-payroll
```

### 7.2 Upload Application Files

**From your local machine:**

```bash
# Upload API files
scp -i ~/private-key.key -r /path/to/oracle-cloud-migration/api/* ubuntu@YOUR_IP:/var/www/hr-payroll/

# Or use git (if you push to GitHub)
ssh -i ~/private-key.key ubuntu@YOUR_IP
cd /var/www/hr-payroll
git clone https://github.com/yourusername/hr-payroll.git .
```

### 7.3 Install Dependencies

```bash
cd /var/www/hr-payroll
npm install
```

### 7.4 Configure Environment

```bash
nano .env
```

Add:
```env
PORT=3000
NODE_ENV=production

DB_HOST=localhost
DB_USER=hrapp
DB_PASSWORD=YourSecurePassword123!
DB_NAME=hrdb

JWT_SECRET=your_random_jwt_secret_here_change_this
ALLOWED_ORIGINS=*
```

Save: `Ctrl+O`, Enter, `Ctrl+X`

### 7.5 Test Application

```bash
node server.js
```

Should see:
```
════════════════════════════════════════════════════
   HR PAYROLL API - Oracle Cloud Migration
════════════════════════════════════════════════════
🚀 Server running on port 3000
✅ MySQL connection pool established
```

Test from your local machine:
```bash
curl http://YOUR_IP:3000/health
```

Should return:
```json
{"status":"healthy","database":"connected"}
```

Press `Ctrl+C` to stop the test server.

---

## Step 8: Configure Auto-Start

### 8.1 Start with PM2

```bash
cd /var/www/hr-payroll
pm2 start server.js --name hr-payroll-api
```

### 8.2 Configure Auto-Restart on Reboot

```bash
pm2 startup
# Copy and run the command it outputs (starts with sudo)

pm2 save
```

### 8.3 Verify PM2 Status

```bash
pm2 status
```

Should show:
```
┌─────┬────────────────────┬─────────────┬─────────┬─────────┬──────────┐
│ id  │ name               │ mode        │ status  │ cpu     │ memory   │
├─────┼────────────────────┼─────────────┼─────────┼─────────┼──────────┤
│ 0   │ hr-payroll-api     │ fork        │ online  │ 0%      │ 45.5mb   │
└─────┴────────────────────┴─────────────┴─────────┴─────────┴──────────┘
```

### 8.4 View Logs

```bash
pm2 logs hr-payroll-api
```

---

## Step 9: Setup SSL (Optional but Recommended)

### 9.1 Install Certbot

```bash
sudo apt install certbot -y
```

### 9.2 Get Free SSL Certificate

**Prerequisites:**
- You need a domain name pointing to your Oracle Cloud IP
- Example: hr.yourdomain.com → 132.145.123.45

```bash
sudo certbot certonly --standalone -d hr.yourdomain.com
```

Follow prompts:
```
Enter email: your@email.com
Agree to terms: Y
Share email: N (your choice)
```

Certificates saved to: `/etc/letsencrypt/live/hr.yourdomain.com/`

### 9.3 Configure HTTPS in Node.js

Update `server.js` to use HTTPS:

```javascript
const https = require('https');
const fs = require('fs');

const options = {
  key: fs.readFileSync('/etc/letsencrypt/live/hr.yourdomain.com/privkey.pem'),
  cert: fs.readFileSync('/etc/letsencrypt/live/hr.yourdomain.com/fullchain.pem')
};

https.createServer(options, app).listen(443, () => {
  console.log('HTTPS server running on port 443');
});
```

### 9.4 Auto-Renew Certificate

```bash
sudo crontab -e
```

Add:
```
0 0 1 * * certbot renew --quiet && pm2 restart hr-payroll-api
```

---

## Step 10: Testing

### 10.1 Test API Endpoints

**Health Check:**
```bash
curl http://YOUR_IP:3000/health
```

**List Employees:**
```bash
curl http://YOUR_IP:3000/api/employees
```

**Get Specific Payslip:**
```bash
curl http://YOUR_IP:3000/api/payslips/1
```

### 10.2 Test from Browser

Open browser:
```
http://YOUR_IP:3000/health
```

Should see JSON response.

### 10.3 Test Frontend

Update your HTML files:
```javascript
const API_BASE_URL = 'http://YOUR_IP:3000';
```

Open HTML file in browser and test:
- Load employees
- Create payslip
- **CRITICAL TEST:** Edit a payslip (should be <1 second, not 2 minutes!)
- View reports

---

## 🎉 Migration Complete!

Your system is now running on Oracle Cloud Always Free tier:

✅ **MySQL database** (50GB storage)
✅ **Node.js API** (4 ARM cores, 24GB RAM)
✅ **No monthly costs** (free forever!)
✅ **Performance:** 2 minutes → <1 second for edits!

---

## Maintenance Commands

**View logs:**
```bash
pm2 logs hr-payroll-api
```

**Restart API:**
```bash
pm2 restart hr-payroll-api
```

**Stop API:**
```bash
pm2 stop hr-payroll-api
```

**Database backup:**
```bash
mysqldump -u hrapp -p hrdb > backup-$(date +%Y%m%d).sql
```

**Update application:**
```bash
cd /var/www/hr-payroll
git pull  # or upload new files
npm install
pm2 restart hr-payroll-api
```

---

## Troubleshooting

### Can't connect to VM via SSH

- Check security list has port 22 open
- Verify using correct private key
- Verify IP address is correct

### Can't access API (port 3000)

```bash
# Check if running
pm2 status

# Check firewall
sudo iptables -L -n | grep 3000

# Check if port is listening
sudo netstat -tulpn | grep 3000
```

### MySQL connection errors

```bash
# Check MySQL running
sudo systemctl status mysql

# Check credentials
mysql -u hrapp -p hrdb

# Check grants
sudo mysql
SHOW GRANTS FOR 'hrapp'@'localhost';
```

### High memory usage

```bash
# Check memory
free -h

# Reduce PM2 instances if needed
pm2 scale hr-payroll-api 1
```

---

## Next Steps

1. ✅ Run data migration script (migrate-data.js)
2. ✅ Update HTML frontend files
3. ✅ Test all functionality
4. ✅ Run parallel with Google Sheets for 1 week
5. ✅ Switch users over gradually
6. ✅ Shut down Apps Script after 1 month

---

## Support

If you encounter issues:

1. Check PM2 logs: `pm2 logs`
2. Check MySQL logs: `sudo tail -f /var/log/mysql/error.log`
3. Check system logs: `sudo journalctl -xe`
4. Test database: `mysql -u hrapp -p hrdb`
5. Test API: `curl http://localhost:3000/health`

---

**🎊 Congratulations! You've successfully migrated to Oracle Cloud for FREE!**
