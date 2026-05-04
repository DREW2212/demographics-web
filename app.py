import io
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, send_file, jsonify
from merger import DemographicsMerger, validate_file_radius, extract_address_from_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

DEFAULT_TEMPLATE = Path(__file__).parent / 'template' / 'Full_Service_Demographic_Template.xlsx'


@app.route('/')
def index():
    return render_template('index.html', has_default_template=DEFAULT_TEMPLATE.exists())


@app.route('/merge', methods=['POST'])
def merge():
    try:
        address = request.form.get('address', '').strip() or None

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Resolve template
            if 'template' in request.files and request.files['template'].filename:
                template_path = tmpdir / 'template.xlsx'
                request.files['template'].save(str(template_path))
            elif DEFAULT_TEMPLATE.exists():
                template_path = DEFAULT_TEMPLATE
            else:
                return jsonify({'error': 'No template file available. Please upload the Full Service Demographic Template.'}), 400

            mile_files = {}

            # ZIP upload
            if 'zip_file' in request.files and request.files['zip_file'].filename:
                zip_path = tmpdir / 'upload.zip'
                request.files['zip_file'].save(str(zip_path))

                with zipfile.ZipFile(zip_path) as zf:
                    for name in zf.namelist():
                        if name.lower().endswith('.xlsx') and not name.startswith('__MACOSX'):
                            radius = validate_file_radius(Path(name).name)
                            if radius:
                                out_path = tmpdir / f'mile_{radius}.xlsx'
                                out_path.write_bytes(zf.read(name))
                                mile_files[radius] = str(out_path)
                                if not address:
                                    address = extract_address_from_filename(Path(name).name)
            else:
                # Individual uploads
                for r in ['1', '3', '5']:
                    f = request.files.get(f'mile_{r}')
                    if f and f.filename:
                        path = tmpdir / f'mile_{r}.xlsx'
                        f.save(str(path))
                        mile_files[r] = str(path)
                        if not address:
                            address = extract_address_from_filename(f.filename)

            missing = [f'{r}-mile' for r in ['1', '3', '5'] if r not in mile_files]
            if missing:
                return jsonify({'error': f'Missing files: {", ".join(missing)}'}), 400

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = tmpdir / f'merged_{timestamp}.xlsx'

            merger = DemographicsMerger(str(template_path))
            success = merger.merge_reports(
                mile_1_path=mile_files['1'],
                mile_3_path=mile_files['3'],
                mile_5_path=mile_files['5'],
                output_path=str(output_path),
                address=address
            )

            if not success:
                return jsonify({'error': 'Merge failed. Check that your files have the expected sheet names.'}), 500

            file_bytes = output_path.read_bytes()

        download_name = f'{address} - Full Service.xlsx' if address else f'merged_{timestamp}.xlsx'

        return send_file(
            io.BytesIO(file_bytes),
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
