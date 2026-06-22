import subprocess
import os
from flask import Flask, Response, send_file
from flask_cors import CORS

app = Flask(__name__)
# Enable CORS so if user opens the HTML file independently via file:// protocol
# the SSE connection to localhost:5000 will succeed.
CORS(app)

@app.route('/')
def index():
    """Serves the main HTML file directly."""
    if os.path.exists('Morocco_Volunteer_Opportunities.html'):
        return send_file('Morocco_Volunteer_Opportunities.html')
    else:
        return """
        <h1>No Data Found</h1>
        <p>The file 'Morocco_Volunteer_Opportunities.html' does not exist yet.</p>
        <p>Run <code>python scrape_europa_opportunities.py</code> at least once to generate it.</p>
        """

@app.route('/update')
def update():
    """
    Runs the scraper script in a subprocess and streams the output 
    to the frontend (HTML template) using Server-Sent Events (SSE).
    """
    def generate():
        # Start the external python script, pipe its output
        # Use -u to force unbuffered stdout so live logs stream correctly
        process = subprocess.Popen(
            ['python', '-u', 'scrape_europa_opportunities.py'],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            encoding='utf-8'
        )
        
        # Stream the stdout line by line
        for line in iter(process.stdout.readline, ''):
            if line:
                # SSE format requires sending data fields followed by double newline
                yield f"data: {line.strip()}\n\n"
            
        process.stdout.close()
        process.wait()
        
    return Response(generate(), mimetype='text/event-stream')

if __name__ == '__main__':
    import sys
    # Force UTF-8 stdout encoding to support printing emojis in Windows cmd/powershell
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except AttributeError:
            pass

    print("🚀 Starting local server...")
    print("👉 Open http://localhost:5000 in your browser.")
    # Binding to 0.0.0.0 will make it available on the local network
    # Disabled reloader to prevent creating buggy duplicate processes on Windows
    app.run(host="127.0.0.1", port=5000, debug=True, use_reloader=False)
