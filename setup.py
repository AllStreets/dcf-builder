from setuptools import setup, find_packages

setup(
    name="dcf-builder",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "xlwings>=0.30.0",
        "yfinance>=0.2.0,<1.0",
        "fredapi>=0.5.0",
        "pandas>=2.0.0",
        "openpyxl>=3.1.0",
        "requests>=2.31.0",
    ],
    extras_require={
        "dev": ["pytest>=7.0.0"],
    },
    python_requires=">=3.8",
    author="Connor Evans",
    description="Python-powered Excel add-in for DCF modeling",
    entry_points={
        "console_scripts": [
            "dcf-builder=dcf_builder.main:main",
        ],
    },
)
