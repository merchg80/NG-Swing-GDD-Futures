"""Build proposed source ZIPs by replacing only reviewed, supplied functions.
No generated code is imported, evaluated, executed, deployed or committed.
"""
import ast,hashlib,io,json,re,textwrap,zipfile
from pathlib import Path


def function_name(code):
    tree=ast.parse(textwrap.dedent(code))
    if len(tree.body)!=1 or not isinstance(tree.body[0],(ast.FunctionDef,ast.AsyncFunctionDef)):
        raise ValueError('Each Python block must contain one complete replacement function, without top-level imports or executable statements.')
    node=tree.body[0]
    if node.decorator_list:raise ValueError('Do not include decorators in replacement blocks; the original decorators are preserved.')
    return node


def build(answer,source_context,root):
    name=source_context.get('file','')
    if name not in ('deal_economics.py','ohio_economics.py','streamlit_app.py'):raise ValueError('No approved application source was supplied.')
    path=Path(root)/name
    if not path.is_file():raise ValueError('The original application source is not available.')
    source=path.read_text();tree=ast.parse(source)
    allowed={}
    for snippet in source_context.get('functions',[]):
        # Supplied snippets may be nested function source; dedent for inspection.
        node=ast.parse(textwrap.dedent(snippet)).body[0]
        allowed[node.name]=snippet
    if not allowed:raise ValueError('No source functions were supplied to the AI. Include Python functions and request a complete replacement function first.')
    blocks=re.findall(r'```python\s*\n(.*?)```',answer,flags=re.S|re.I)
    if not blocks:raise ValueError('No complete Python function block was returned. Ask for one complete replacement function per Python block.')
    edits=[];seen=set();lines=source.splitlines(keepends=True)
    for block in blocks:
        replacement=function_name(block);key=replacement.name
        if key not in allowed:raise ValueError('The proposed function was not in the supplied source context: '+key)
        if key in seen:raise ValueError('More than one replacement was returned for '+key)
        seen.add(key)
        matches=[n for n in ast.walk(tree) if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef)) and n.name==key and ast.get_source_segment(source,n)==allowed[key]]
        if len(matches)!=1:raise ValueError('Original function changed or is ambiguous. Prepare a new AI request: '+key)
        original=matches[0]
        if type(original)!=type(replacement) or ast.dump(original.args)!=ast.dump(replacement.args):raise ValueError('The function signature changed. This requires a wider caller review before packaging: '+key)
        # Replace def/body only; retain existing decorators and surrounding code.
        indent=' '*original.col_offset
        body=textwrap.dedent(block).strip('\n')
        new='\n'.join(indent+line if line.strip() else '' for line in body.splitlines())+'\n'
        edits.append((original.lineno-1,original.end_lineno,new,key))
    for start,end,new,key in sorted(edits,reverse=True):lines[start:end]=[new]
    updated=''.join(lines)
    ast.parse(updated)
    compile(updated,name,'exec')  # Compile for syntax/scope checks only; never execute.
    metadata=dict(file=name,original_sha256=hashlib.sha256(source.encode()).hexdigest(),proposed_sha256=hashlib.sha256(updated.encode()).hexdigest(),functions=sorted(seen),validation='Syntax and replacement scope/signature checked; runtime behavior untested.')
    readme='PROPOSED AI CODE UPDATE - REVIEW BEFORE UPLOADING\n\nFile: '+name+'\nFunctions: '+', '.join(sorted(seen))+'\n\nThe Python file is a full replacement built from this app\'s current source, with only the listed function bodies replaced. Existing decorators and surrounding source are preserved. Compare against GitHub before uploading; another user may have made a newer change. Original SHA-256 is recorded in change_manifest.json.\n\nNo generated code was executed or deployed. Syntax checks are not behavioral or security tests. Review the proposed functions, validate calculations and test in a separate branch/app before replacing production. Keep a rollback copy. No real API keys or app Secrets are included. The unchanged app source may contain embedded business defaults.\n\n'+answer
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr(name,updated);z.writestr('change_manifest.json',json.dumps(metadata,indent=2));z.writestr('REVIEW_BEFORE_UPLOAD.txt',readme)
    return out.getvalue(),metadata
