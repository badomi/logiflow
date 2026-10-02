from pathlib import Path
from datetime import datetime
import zipfile,sqlite3
root=Path('C:/dev/logiflow-main'); dest=Path('C:/dev/backups')/('merge-'+datetime.now().strftime('%Y%m%d-%H%M%S')); dest.mkdir(parents=True)
with zipfile.ZipFile(dest/'project.zip','w',zipfile.ZIP_DEFLATED) as z:
 for p in root.rglob('*'):
  if not p.is_file() or any(x in p.parts for x in ('node_modules','.venv','__pycache__','.pytest_cache')) or p.suffix in ('.db','.db-wal','.db-shm'): continue
  z.write(p,p.relative_to(root))
src=sqlite3.connect('file:'+str(root/'server/data/leona.db')+'?mode=ro',uri=True); target=sqlite3.connect(dest/'leona.db'); src.backup(target); target.close(); src.close()
print(dest)
