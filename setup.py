import os
from setuptools import setup
from io import open

here = os.path.abspath(os.path.dirname(__file__))

about = {}
with open(os.path.join(here, "pptx_tables", "__version__.py"), "r", encoding="utf-8") as f:
    exec(f.read(), about)

with open("README.rst", "r", encoding="utf-8") as f:
    readme = f.read()

with open("HISTORY.rst", "r", encoding="utf-8") as f:
    history = f.read()

setup(name=about["__title__"],
      version=about["__version__"],
      description=about["__description__"],
      long_description=readme + "\n\n" + history,
      author=about["__author__"],
      author_email=about["__author_email__"],
      url=about["__url__"],
      classifiers=["Programming Language :: Python :: 3.6",
                   "Intended Audience :: Developers",
                   "Development Status :: 2 - Pre-Alpha"],
      platforms=["Operating System :: MacOS :: MacOS X"],
      packages=[about["__title__"]],
      install_requires=["python-pptx"],
      )
