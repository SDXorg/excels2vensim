from setuptools import setup, find_packages

exec(open('excels2vensim/_version.py').read())

setup(
    name='excels2vensim',
    version=__version__,
    python_requires='>=3.9',
    author='Eneko Martin Martinez',
    author_email='eneko.martin.martinez@gmail.com',
    packages=find_packages(exclude=['docs', 'tests', 'dist', 'build']),
    url='https://github.com/SDXorg/excels2vensim',
    license='MIT',
    description='Easy write excel inputs in Vensim',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    keywords=['System Dynamics', 'Vensim', 'Excel'],
    classifiers=[
        'Development Status :: 3 - Alpha',

        'Topic :: Scientific/Engineering :: Information Analysis',
        'Intended Audience :: Science/Research',

        'License :: OSI Approved :: MIT License',

        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
    ],
    install_requires=open('requirements.txt').read().strip().split('\n'),
    include_package_data=True
)
