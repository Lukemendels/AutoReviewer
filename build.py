import os

def build():
    workspace = '/home/luke/AutoReviewer'
    template_path = os.path.join(workspace, 'src', 'autoreviewer.template.html')
    jszip_path = os.path.join(workspace, 'src', 'jszip.min.js')
    output_path = os.path.join(workspace, 'autoreviewer.html')
    
    print("Reading JSZip source code...")
    with open(jszip_path, 'r', encoding='utf-8') as f:
        jszip_content = f.read()
        
    print("Reading HTML template...")
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
        
    print("Bundling JSZip into HTML...")
    # Replace the placeholder with the JSZip minified content
    bundled_content = template_content.replace('/* {{JSZIP_MIN_JS}} */', jszip_content)
    
    print("Writing finalized HTML...")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(bundled_content)
        
    print(f"Build complete! Saved to {output_path}")

if __name__ == '__main__':
    build()
