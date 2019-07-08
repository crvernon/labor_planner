"""setup.py

Module wrapper for `labor_planner`

@author: Chris R. Vernon

"""


class VersionError(Exception):
    def __init__(self, *args, **kwargs):
        Exception.__init__(self, *args, **kwargs)


try:
    from setuptools import setup, find_packages
except ImportError:
    print("Must have setuptools installed to run setup.py. Please install and try again.")
    raise


def readme():
    with open('README.md') as f:
        return f.read()


def get_requirements():
    with open('requirements.txt') as f:
        return f.read().split()

setup(
    name='labor_planner',
    version='1.0.0',
    packages=find_packages(),
    url='https://github.com/crvernon/labor_planner',
    license='MIT',
    author='Chris R. Vernon',
    author_email='chris.vernon@pnnl.gov',
    description='Plan and visualize staff labor calendar year allocation',
    python_requires='>=3.3.*, <4'
)
