from setuptools import find_packages, setup

setup(
    name = "Tech_daily_reports",
    version = "0.0.1",
    install_requires = [
        'pandas',
        'datetime',
        'openpyxl',
        'jinja2,'
        'importlib-metadata; python_version < "3.12"'
        ],
    packages = find_packages()
)
