"""Assemble index.html from src/head.html + src/body.html, injecting data/data.json and src/logo.png."""
import base64, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
logo = 'data:image/png;base64,' + base64.b64encode((root / 'src/logo.png').read_bytes()).decode()
html = ((root / 'src/head.html').read_text() + (root / 'src/body.html').read_text()).replace('__LOGO__', logo).replace('__DATA__', (root / 'data/data.json').read_text())
(root / 'index.html').write_text(html); print('wrote index.html')
