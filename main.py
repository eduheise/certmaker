from glob import glob
from certmaker import CertMaker

if __name__ == '__main__':
    templates = glob('templates/Certificate*')
    for template in templates:
        cert = CertMaker(template)
        cert.generate()
