import subprocess


def convert_to_pdf(name):
    str = "libreoffice6.3 --convert-to pdf:writer_pdf_Export ./pdf/{name}.xls --outdir ./pdf/".format(name=name)
    subprocess.call(str, shell=True)


if __name__=="__main__":
    convert_to_pdf("160301010124")
