from setuptools import setup, find_packages

setup(
    name="email-harvester",
    version="1.0.0",
    description="Private CLI tool for scraping US business directories and extracting contact emails",
    python_requires=">=3.11",
    packages=find_packages(),
    install_requires=open("requirements.txt").read().splitlines(),
    entry_points={
        "console_scripts": [
            "email-harvester=email_harvester.cli:main",
        ],
    },
)
