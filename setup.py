import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="aa-py-openpyxl-util",
    author="Rudolf Byker",
    author_email="rudolfbyker@gmail.com",
    description="Utilities that build on top of `openpyxl`.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/AutoActuary/aa-py-openpyxl-util",
    packages=setuptools.find_packages(exclude=["test"]),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: Other/Proprietary License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
    use_scm_version={
        "write_to": "aa_py_openpyxl_util/version.py",
    },
    setup_requires=[
        "setuptools_scm",
    ],
    install_requires=[
        "locate>=1.1.1,==1.*",
        "openpyxl>=3.1.0,==3.1.*",  # openpyxl does not use semantic versioning.
        "aa_py_xl_convert==0.0.1",
    ],
    package_data={
        "": [
            "py.typed",
        ],
    },
)
