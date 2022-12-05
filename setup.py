import setuptools

with open('README.md', 'r') as fp:
    long_description = fp.read()

setuptools.setup(
    name='easy_generate_excel',
    version="0.0.5",
    author='Daniil y0gur-dv',
    description='Easy generate excel files',
    packages=['easy_generate_excel'],
    install_requires=[
        'openpyxl',
    ],
    url='https://github.com/y0gurt-dv/Easy-generate-excel',
    keywords=['excel', 'generators', 'easy'],
    long_description=long_description,
    long_description_content_type='text/markdown',
)
