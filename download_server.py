#!/usr/bin/env python3
"""Simple HTTP server to serve the Excel file for download."""
import http.server
import os

os.chdir("/home/user/Excel")

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/" or self.path == "":
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write(b"""
            <html>
            <head><title>Download Home Renovation Management System</title>
            <style>
                body { font-family: Arial, sans-serif; display: flex; justify-content: center;
                       align-items: center; min-height: 100vh; margin: 0;
                       background: linear-gradient(135deg, #1F3864, #2E75B6); color: white; }
                .card { background: white; color: #333; border-radius: 16px; padding: 48px;
                        text-align: center; box-shadow: 0 8px 32px rgba(0,0,0,0.3); max-width: 500px; }
                h1 { color: #1F3864; margin-bottom: 8px; }
                p { color: #666; margin-bottom: 24px; }
                a.btn { display: inline-block; background: #2E75B6; color: white; padding: 16px 40px;
                        border-radius: 8px; text-decoration: none; font-size: 18px; font-weight: bold;
                        transition: background 0.3s; }
                a.btn:hover { background: #1F3864; }
                .info { margin-top: 24px; font-size: 14px; color: #999; }
            </style></head>
            <body>
                <div class="card">
                    <h1>&#127968; Home Renovation Management System</h1>
                    <p>Comprehensive Excel Workbook &bull; 56 Sheets &bull; 7.0 MB</p>
                    <a class="btn" href="/Home_Renovation_Management_System.xlsx" download>
                        &#11015; Download Excel File
                    </a>
                    <div class="info">
                        <p>Python script also available: <a href="/renovation_management_system.py" download>renovation_management_system.py</a></p>
                    </div>
                </div>
            </body>
            </html>
            """)
        else:
            super().do_GET()

print("Server running on port 8080...")
http.server.HTTPServer(("0.0.0.0", 8080), Handler).serve_forever()
