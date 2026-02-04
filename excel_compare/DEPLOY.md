# Deploy Excel Comparison Tool on a VPS

Steps to run the app on a Linux VPS (Ubuntu/Debian).

---

## If you already have another Flask app on this VPS

You can run **both** on the same server. Each app needs:

| | Existing Flask app | Excel Compare app |
|---|-------------------|-------------------|
| **Folder** | e.g. `/var/www/myapp` | e.g. `/var/www/excel_compare` |
| **Port** | e.g. `5000` (or 8000) | **Different port**, e.g. `5001` |
| **Virtualenv** | Its own `venv` | Its own `venv` (separate) |
| **systemd service** | e.g. `myapp.service` | `excel-compare.service` |

**Important:** Use a **different port** for Excel Compare so it doesn’t clash with the other Flask app. Example: existing app on `127.0.0.1:5000`, Excel Compare on `127.0.0.1:5001`.

**Nginx:** Add a second `server` block (subdomain) or another `location` (path) that proxies to the Excel Compare port:

- **Subdomain:** e.g. `excel.yourdomain.com` → `proxy_pass http://127.0.0.1:5001;`
- **Path:** e.g. `yourdomain.com/excel/` → `proxy_pass http://127.0.0.1:5001/;` (and run this app with a URL prefix if needed).

Then in the sections below, use **port 5001** (or another free port) for Excel Compare instead of 5000.

---

## 1. Prepare the VPS

- SSH into your VPS.
- Update and install Python 3 and venv:

```bash
sudo apt update
sudo apt install -y python3 python3-pip python3-venv
```

## 2. Upload the project

Copy the `excel_compare` folder to the server (e.g. `/var/www/excel_compare` or `~/excel_compare`).

Options:

- **Git:** `git clone <your-repo-url>` then `cd excel_compare`
- **rsync** (from your PC):  
  `rsync -avz excel_compare/ user@your-vps-ip:/var/www/excel_compare/`
- **SCP:**  
  `scp -r excel_compare user@your-vps-ip:/var/www/`

## 3. Create virtualenv and install dependencies

On the VPS:

```bash
cd /var/www/excel_compare   # or your path
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

## 4. Run with Gunicorn (production)

From the **excel_compare** directory (with `venv` activated):

```bash
gunicorn --bind 0.0.0.0:5000 --workers 2 --threads 2 app:app
```

- `0.0.0.0:5000` — listen on all interfaces, port 5000.
- `app:app` — module `app`, Flask `app` object.
- Adjust `--workers` as needed (e.g. 2–4 for a small VPS).

Test: open `http://YOUR_VPS_IP:5000` in a browser.

## 5. Run as a systemd service (recommended)

So the app restarts on reboot and you can manage it with `systemctl`:

1. Create a service file:

```bash
sudo nano /etc/systemd/system/excel-compare.service
```

2. Paste (adjust paths if different):

```ini
[Unit]
Description=Excel Comparison Tool
After=network.target

[Service]
User=www-data
Group=www-data
WorkingDirectory=/var/www/excel_compare
Environment="PATH=/var/www/excel_compare/venv/bin"
ExecStart=/var/www/excel_compare/venv/bin/gunicorn --bind 127.0.0.1:5000 --workers 2 --threads 2 app:app
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

- If your app is in a different directory or under your user, change `User`, `Group`, `WorkingDirectory`, and paths in `Environment` and `ExecStart`.
- `127.0.0.1:5000` means only localhost can reach the app; use Nginx in front (see below) for public access.

3. Enable and start:

```bash
sudo systemctl daemon-reload
sudo systemctl enable excel-compare
sudo systemctl start excel-compare
sudo systemctl status excel-compare
```

Logs: `sudo journalctl -u excel-compare -f`

## 6. (Optional) Nginx in front

Use Nginx as reverse proxy and (optionally) HTTPS.

1. Install Nginx (if not already):

```bash
sudo apt install -y nginx
```

2. Add a site config:

```bash
sudo nano /etc/nginx/sites-available/excel-compare
```

**If this is your only app** (replace `YOUR_DOMAIN_OR_IP`):

```nginx
server {
    listen 80;
    server_name YOUR_DOMAIN_OR_IP;

    client_max_body_size 32M;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

**If you already have another Flask app** (e.g. main site on port 5000), use a **different port** for Excel Compare (e.g. 5001) in the systemd unit, then add one of these:

- **Subdomain** (e.g. `excel.yourdomain.com` → Excel Compare):

```nginx
server {
    listen 80;
    server_name excel.yourdomain.com;

    client_max_body_size 32M;

    location / {
        proxy_pass http://127.0.0.1:5001;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

- **Path under same domain** (e.g. `yourdomain.com/excel/` → Excel Compare). In the **existing** server block for your domain, add:

```nginx
    location /excel/ {
        proxy_pass http://127.0.0.1:5001/;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        client_max_body_size 32M;
    }
```

(Excel Compare app must be configured to run under `/excel` prefix, or use the trailing slash so Nginx strips `/excel` and forwards `/` to the app.)

3. Enable and reload:

```bash
sudo ln -s /etc/nginx/sites-available/excel-compare /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

Open `http://YOUR_DOMAIN_OR_IP`. For HTTPS, use Certbot: `sudo apt install certbot python3-certbot-nginx && sudo certbot --nginx -d YOUR_DOMAIN`.

## Quick checklist

| Step              | Command / action                                      |
|-------------------|--------------------------------------------------------|
| Install Python    | `sudo apt install python3 python3-pip python3-venv`   |
| Upload project    | rsync / scp / git clone                               |
| Venv + deps       | `python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt` |
| Run once          | `gunicorn --bind 0.0.0.0:5000 app:app`                |
| Run as service    | systemd unit above, then `systemctl enable --now excel-compare` |
| Public + HTTPS    | Nginx reverse proxy + Certbot                         |

## Firewall

If you use UFW and only Nginx is public:

```bash
sudo ufw allow 80
sudo ufw allow 443
sudo ufw allow 22
sudo ufw enable
```

Do **not** open port 5000 if the app is bound to `127.0.0.1` and only Nginx talks to it.

---

## Common doubts (VPS already has Flask)

**Q: Will this overwrite or break my existing Flask app?**  
No. This app lives in its **own folder** with its **own venv** and runs on a **different port** (e.g. 5001). The other app keeps using its port (e.g. 5000).

**Q: Do I need a separate Python/venv?**  
Yes. Use a **separate virtualenv** inside `excel_compare` so dependencies (Flask, pandas, etc.) don’t conflict with the other project.

**Q: How does Nginx know which app to use?**  
By **server_name** (subdomain) or **location** (path). For example: `excel.yourdomain.com` or `yourdomain.com/excel/` goes to Excel Compare; the rest goes to your other app.

**Q: What port should Excel Compare use?**  
Any free port. If the existing Flask app uses 5000, use **5001** (or 8000, 8080, etc.) for Excel Compare. In the systemd unit, use:  
`--bind 127.0.0.1:5001`
