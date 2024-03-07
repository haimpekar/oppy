'''
TODO:
1. install flask-limiter
'''
import os
import subprocess
from flask import Flask, request, send_file, make_response, render_template_string
from datetime import datetime
from werkzeug.utils import secure_filename
from sys import modules

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'xlsx'}
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

process = None
filepath = None
# max_stream_size = 100000


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


index_html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Upload</title>
</head>
<body>
    <h1>File Upload</h1>
    <form id="uploadForm" enctype="multipart/form-data">
        <label for="fileInput">Select a file:</label>
        <input type="file" id="fileInput" name="fileInput" accept=".xlsx">
        <br>
        <input type="button" value="Upload File" onclick="uploadFile()">
    </form>
    <div id="message"></div>

    <script>
        function uploadFile() {
            var fileInput = document.getElementById('fileInput');
            var file = fileInput.files[0];
            if (!file) {
                document.getElementById('message').innerText = 'Please select a file.';
                return;
            }
            var fileName = file.name;
            if (!fileName.endsWith('.xlsx')) {
                document.getElementById('message').innerText = 'Please select an .xlsx file.';
                return;
            }
            var formData = new FormData();
            formData.append('file', file);

            fetch('/', {
                method: 'POST',
                body: formData
            })
            .then(response => response.text())
            .then(message => {
                document.getElementById('message').innerText = message;
            })
            .catch(error => {
                console.error('Error:', error);
            });
        }
        
        // Prevent default form submission
        document.getElementById('uploadForm').addEventListener('submit', function(event) {
            event.preventDefault();
        });
    </script>
</body>
</html>
"""


@app.route('/', methods=['POST', 'GET'])
def index():
    if request.method == 'POST':
        # Log basic request information
        try:
            uploaded_file = request.files['file']
            if uploaded_file:
                if allowed_file(uploaded_file.filename):
                    file_flag = True
                else:
                    return make_response('Invalid file or file extension', {'Sender': 'Python', 'Status': 'Rejected'})
            else:
                file_flag = False
        except:
            file_flag = False
        try:
            log_data = {
                "timestamp": datetime.utcnow().isoformat(),
                "remote_addr": request.remote_addr,
                "url": request.url,
                "sender": request.headers["sender"],
                "content": request.headers["content"],
                "data": request.data
            }
            print(log_data)
        except:
            print(request)
        if not file_flag:
            if ('Sender' not in request.headers) and ('Content' not in request.headers):
                return make_response('Wrong header', {'Sender': 'Python', 'Status': 'Rejected'})
            elif (request.headers['Content'] != 'Stream') and (request.headers['Content'] != 'Request'):
                return make_response('Wrong content', {'Sender': 'Python', 'Status': 'Rejected'})
            elif request.headers['Sender'] != 'VBA':
                return make_response('Wrong sender', {'Sender': 'Python', 'Status': 'Rejected'})
            elif request.content_length > 100000:
                return make_response('Wrong stream size', {'Sender': 'Python', 'Status': 'Rejected'})
        global process, filepath
        if process is None:
            process_flag = 'Waiting'
        elif process.poll() is None:
            process_flag = 'Processing'
        else:
            process_flag = 'Completed'
        if file_flag or ((request.headers['Content'] == 'Stream') and (process_flag != 'Processing')):
            try:
                if file_flag:
                    filename = secure_filename(uploaded_file.filename)
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    request.files['file'].save(filepath)
                else:
                    filepath = request.data
                # Process the file
                process = subprocess.Popen(['python', os.getcwd() + '\\oppy.py', filepath, "server"], stdout=subprocess.PIPE, text=True)
                # Send back acknowledgment message
                return make_response('Stream uploaded successfully', {'Sender': 'Python', 'Status': 'Accepted'})
                # return f'Stream uploaded successfully'
            except Exception as e:
                return make_response(f'Error processing stream: {e}', {'Sender': 'Python', 'Status': 'Error'})
        elif process_flag == 'Waiting':
            return make_response('No active process', {'Sender': 'Python', 'Status': 'Rejected'})
        else:
            try:
                # first get rid of all the messages one by one
                for line in process.stdout:
                    return make_response(line.strip(), {'Sender': 'Python', 'Status': 'Processing'})
                if file_flag:
                    return send_file(filepath, as_attachment=True)
                else:
                    return make_response('Mission Completed', {'Sender': 'Python', 'Status': 'Completed'})
            except Exception as e:
                return make_response(f'Error processing messages: {e}', {'Sender': 'Python', 'Status': 'Error'})
    else:
        return render_template_string(index_html)


if __name__ == '__main__':
    if "pydevd" in modules:
        app.run(host='0.0.0.0', port=5000, debug=False)
        # app.run(debug=False)
    else:
        app.run(host='0.0.0.0', port=5000, debug=False)
