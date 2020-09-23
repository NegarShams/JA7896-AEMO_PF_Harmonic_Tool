import py_compile
import os
from pathlib import Path


current_dir = Path(__file__).parent
target_dir = Path(current_dir).joinpath('compiled')

if __name__ == '__main__':
	# check if file already exists
	if not Path(target_dir).is_dir():
		Path(target_dir).mkdir()

	for file in Path(current_dir).rglob('*.*'):
		file_name = file.name
		# skip test_files
		source_dir = Path(file).parent
		exclude = ('test', '.git', '.idea', 'hooks', '__pycache__')
		skip = any([any([x in file_name for x in exclude]), any([x in str(source_dir) for x in exclude])])
		if skip:
			continue
		target = Path(str(source_dir).replace(str(current_dir), str(target_dir)))
		if not target.is_dir():
			target.mkdir()

		target_file = target.joinpath(file_name+'c')
		print(file)
		print(target_file)

		if str(file).endswith('.py'):
			compiled = py_compile.compile(file=file, cfile=target_file)
			print(compiled)
		print('')