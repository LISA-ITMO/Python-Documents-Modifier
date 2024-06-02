# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

import os
import sys
for x in os.walk('/Users/vladtereshch/PycharmProjects/PythonDM/src'):
  sys.path.insert(0, x[0])
sys.path.insert(0, os.path.abspath('..'))

site_packages = r'~/PycharmProjects/PythonDM/venv/lib/python3.9/site-packages'
if site_packages not in sys.path:
    sys.path.append(site_packages)

project = 'Edulytica-PythonDM'
copyright = '2024, Edulytica Team'
author = 'Edulytica Team'
release = '0.0.1'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = ['sphinx.ext.todo', 'sphinx.ext.viewcode', 'sphinx.ext.autodoc']
todo_include_todos = True

templates_path = ['_templates']
exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'alabaster'
html_static_path = ['_static']
