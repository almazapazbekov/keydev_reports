from setuptools import setup, find_packages

# def readme():
#   with open('README.md', 'r') as f:
#     return f.read()

setup(
    name='keydev_reports',
    version='1.0.0',
    author='KeyDev',
    author_email='almazapazbekov@gmail.com',
    description='генерация отчетов',
    # long_description=readme(),
    # long_description_content_type='text/markdown',
    url='https://github.com/KeyDev-LTD/key_dev_reports',
    packages=['keydev_reports'],
    install_requires=['python-docx==0.8.11',
                      'python-docx-replace==0.4.1',
                      'openpyxl==3.0.10',
],
    classifiers=[
        'Programming Language :: Python :: 3.11',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent'
    ],
    keywords='example python',
    project_urls={
        'Documentation': 'link'
    },
    python_requires='>=3.10'
)
