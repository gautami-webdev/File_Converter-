from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from pathlib import Path
from io import BytesIO
import tempfile, os, csv, json, ast, subprocess, shutil

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024
INPUTS = {'csv','py','xlsx','xlsm','json','txt','md','docx','pptx','pdf','png','jpg','jpeg'}
OUTPUTS = ['csv','py','xlsx','json','txt','md','pdf','docx','pptx']
MIME = {'csv':'text/csv; charset=utf-8','py':'text/x-python; charset=utf-8','xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','json':'application/json; charset=utf-8','txt':'text/plain; charset=utf-8','md':'text/markdown; charset=utf-8','pdf':'application/pdf','docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','pptx':'application/vnd.openxmlformats-officedocument.presentationml.presentation'}

def read_csv(path):
    with open(path, encoding='utf-8-sig', newline='') as f:
        sample=f.read(8192); f.seek(0)
        try: dialect=csv.Sniffer().sniff(sample, delimiters=',;\t|')
        except csv.Error: dialect=csv.excel
        reader=csv.DictReader(f, dialect=dialect)
        if not reader.fieldnames: raise ValueError('CSV must contain a header row.')
        return list(reader), list(reader.fieldnames)

def csv_bytes(rows, headers):
    out=BytesIO(); import io
    wrapper=io.TextIOWrapper(out, encoding='utf-8-sig', newline='', write_through=True)
    writer=csv.DictWriter(wrapper, fieldnames=headers, extrasaction='ignore'); writer.writeheader(); writer.writerows(rows); wrapper.flush()
    return out.getvalue()

def safe_py(path):
    try: tree=ast.parse(Path(path).read_text(encoding='utf-8-sig'))
    except (SyntaxError,UnicodeDecodeError) as e: raise ValueError('The Python file is not valid UTF-8 Python syntax.') from e
    if len(tree.body)!=1 or not isinstance(tree.body[0],ast.Assign) or len(tree.body[0].targets)!=1 or not isinstance(tree.body[0].targets[0],ast.Name):
        raise ValueError("PY → CSV accepts one literal assignment, e.g. data = [{'name': 'A'}]. Uploaded code is never executed.")
    try: value=ast.literal_eval(tree.body[0].value)
    except (ValueError,TypeError,SyntaxError) as e: raise ValueError('The assigned value must be a literal list/dictionary, not an expression.') from e
    if isinstance(value,dict):
        if value and all(isinstance(v,list) for v in value.values()):
            if len({len(v) for v in value.values()})>1: raise ValueError('All column lists must have equal lengths.')
            headers=list(value); rows=[{h:value[h][i] for h in headers} for i in range(len(next(iter(value.values()),[])))]
        else: headers=list(value); rows=[value]
    elif isinstance(value,list) and value and all(isinstance(x,dict) for x in value):
        headers=list(dict.fromkeys(k for x in value for k in x)); rows=[{h:x.get(h,'') for h in headers} for x in value]
    else: raise ValueError('Use a non-empty list of dictionaries or a dictionary of equal-length lists.')
    for row in rows:
        for k,v in list(row.items()): row[k]='' if v is None else json.dumps(v,ensure_ascii=False) if isinstance(v,(dict,list,tuple)) else str(v)
    return rows,headers

def text_content(path, ext):
    if ext in {'txt','md','py','csv','json'}: return Path(path).read_text(encoding='utf-8-sig')
    if ext=='pdf':
        from pypdf import PdfReader
        return '\n\n'.join(p.extract_text() or '' for p in PdfReader(path).pages)
    if ext=='docx':
        from docx import Document
        d=Document(path); chunks=[p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows: chunks.append('\t'.join(c.text for c in r.cells))
        return '\n'.join(chunks)
    if ext=='pptx':
        from pptx import Presentation
        chunks=[]
        for i,s in enumerate(Presentation(path).slides,1):
            chunks.append(f'Slide {i}')
            for sh in s.shapes:
                if getattr(sh,'has_text_frame',False): chunks.append(sh.text)
                if getattr(sh,'has_table',False):
                    for r in sh.table.rows: chunks.append('\t'.join(c.text for c in r.cells))
        return '\n'.join(chunks)
    raise ValueError('Text extraction is not supported for this file.')

def make_pdf(text):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from xml.sax.saxutils import escape
    out=BytesIO(); doc=SimpleDocTemplate(out,pagesize=A4,leftMargin=18*mm,rightMargin=18*mm,topMargin=18*mm,bottomMargin=18*mm); styles=getSampleStyleSheet(); story=[]
    for line in text.splitlines() or ['']:
        story.extend([Paragraph(escape(line) if line else ' ',styles['BodyText']),Spacer(1,3)])
    doc.build(story); return out.getvalue()

def office_pdf(path, temp):
    exe=shutil.which('libreoffice') or shutil.which('soffice')
    if not exe: raise RuntimeError('Office-to-PDF needs LibreOffice installed and available on PATH.')
    try: subprocess.run([exe,'--headless','--convert-to','pdf','--outdir',temp,path],capture_output=True,text=True,timeout=45,check=True)
    except Exception as e: raise RuntimeError('LibreOffice could not convert this Office file to PDF.') from e
    out=os.path.join(temp,Path(path).stem+'.pdf')
    if not os.path.isfile(out): raise RuntimeError('PDF output was not created. Check LibreOffice installation.')
    return Path(out).read_bytes()

def convert(path, src, dst, temp):
    stem=Path(path).stem[:100]
    if src=='csv':
        rows,headers=read_csv(path)
        if dst=='py':
            from pprint import pformat
            data=('# Generated from CSV; data only, no code execution.\ndata = '+pformat([{h:r.get(h,'') for h in headers} for r in rows],width=100,sort_dicts=False)+'\ncolumns = '+pformat(headers)+'\n').encode()
        elif dst=='json': data=json.dumps(rows,ensure_ascii=False,indent=2).encode()
        elif dst=='xlsx':
            import pandas as pd
            out=BytesIO()
            with pd.ExcelWriter(out,engine='openpyxl') as w: pd.DataFrame(rows,columns=headers).to_excel(w,index=False,sheet_name='Data')
            data=out.getvalue()
        elif dst=='txt': data=Path(path).read_bytes()
        else: raise ValueError(f'CSV → {dst.upper()} is not supported.')
    elif src=='py' and dst=='csv':
        rows,headers=safe_py(path); data=csv_bytes(rows,headers)
    elif src in {'xlsx','xlsm'}:
        import pandas as pd
        df=pd.read_excel(path,dtype=str).fillna('')
        if dst=='csv': data=df.to_csv(index=False).encode('utf-8-sig')
        elif dst=='json': data=df.to_json(orient='records',force_ascii=False,indent=2).encode()
        elif dst=='txt': data=df.to_csv(index=False,sep='\t').encode()
        elif dst=='pdf': data=office_pdf(path,temp)
        else: raise ValueError(f'Excel → {dst.upper()} is not supported.')
    elif src=='json':
        value=json.loads(Path(path).read_text(encoding='utf-8-sig'))
        if dst=='py':
            from pprint import pformat
            data=('# Generated data only.\ndata = '+pformat(value,width=100,sort_dicts=False)+'\n').encode()
        elif dst=='txt': data=Path(path).read_bytes()
        elif dst=='csv':
            if isinstance(value,dict) and value and all(isinstance(v,list) for v in value.values()):
                if len({len(v) for v in value.values()})>1: raise ValueError('Column lists in JSON must have equal lengths.')
                headers=list(value); rows=[{h:value[h][i] for h in headers} for i in range(len(next(iter(value.values()),[])))]
            elif isinstance(value,dict): headers=list(value); rows=[value]
            elif isinstance(value,list) and all(isinstance(x,dict) for x in value): headers=list(dict.fromkeys(k for x in value for k in x)); rows=[{h:x.get(h,'') for h in headers} for x in value]
            else: raise ValueError('JSON must be a list of objects or a dictionary.')
            rows=[{k:'' if v is None else json.dumps(v,ensure_ascii=False) if isinstance(v,(list,dict)) else str(v) for k,v in r.items()} for r in rows]; data=csv_bytes(rows,headers)
        else: raise ValueError(f'JSON → {dst.upper()} is not supported.')
    elif src in {'txt','md','docx','pptx','pdf'}:
        if src=='pdf' and dst=='pdf': raise ValueError('Input is already PDF.')
        text=text_content(path,src)
        if dst=='txt': data=text.encode()
        elif dst=='md': data=text.encode()
        elif dst=='json' and src in {'txt','md'}: data=json.dumps({'text':text},ensure_ascii=False,indent=2).encode()
        elif dst=='pdf': data=office_pdf(path,temp) if src in {'docx','pptx'} else make_pdf(text)
        elif dst=='docx':
            from docx import Document
            d=Document()
            for line in text.splitlines(): d.add_paragraph(line)
            out=BytesIO(); d.save(out); data=out.getvalue()
        elif dst=='pptx':
            from pptx import Presentation
            prs=Presentation(); lines=text.splitlines() or ['']; chunks=[lines[i:i+12] for i in range(0,len(lines),12)]
            for i,chunk in enumerate(chunks):
                s=prs.slides.add_slide(prs.slide_layouts[1]); s.shapes.title.text='Converted Text' if i==0 else f'Converted Text (continued {i+1})'; s.placeholders[1].text='\n'.join(chunk)[:5000]
            out=BytesIO(); prs.save(out); data=out.getvalue()
        else: raise ValueError(f'.{src} → {dst.upper()} is not supported.')
    elif src in {'png','jpg','jpeg'} and dst=='pdf':
        from PIL import Image
        with Image.open(path) as im:
            if im.mode not in ('RGB','L'): im=im.convert('RGB')
            out=BytesIO(); im.save(out,format='PDF',resolution=150); data=out.getvalue()
    else: raise ValueError(f'Conversion from .{src} to .{dst} is not supported.')
    suffix=dst
    mime=MIME.get(dst,'application/octet-stream')
    return data,f'{stem}.{suffix}',mime

@app.get('/')
def home(): return render_template('index.html', outputs=OUTPUTS)
@app.post('/api/convert')
def api_convert():
    up=request.files.get('file'); dst=(request.form.get('target') or '').lower().lstrip('.')
    if not up or not up.filename: return jsonify(error='Please choose a file first.'),400
    name=secure_filename(up.filename); src=Path(name).suffix.lower().lstrip('.')
    if not name or src not in INPUTS: return jsonify(error='Unsupported file type. Please select a supported file.'),400
    if dst not in OUTPUTS: return jsonify(error='Unsupported output format.'),400
    if src==dst: return jsonify(error='Choose an output format different from the input.'),400
    try:
        with tempfile.TemporaryDirectory(prefix='fileflow_') as temp:
            path=os.path.join(temp,name); up.save(path)
            data,outname,mime=convert(path,src,dst,temp)
        return send_file(BytesIO(data),mimetype=mime,as_attachment=True,download_name=outname,max_age=0)
    except ValueError as e: return jsonify(error=str(e)),400
    except RuntimeError as e: return jsonify(error=str(e)),422
    except Exception:
        app.logger.exception('Conversion failed')
        return jsonify(error='Conversion failed. Check that the file is valid and not password-protected.'),500
@app.errorhandler(413)
def too_large(_): return jsonify(error='File too large. Maximum size is 25 MB.'),413
if __name__=='__main__': app.run(debug=True,port=int(os.environ.get('PORT','5000')))
