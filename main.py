import click
from win32com import client as com
from win32com.client import constants
from pywintypes import com_error
from collections import OrderedDict
from pkg_resources import iter_entry_points

try:
    Word = com.gencache.EnsureDispatch('Word.Application')
except com_error as e:
    raise click.ClickException(e.excepinfo[2])
except Exception as e:
    raise click.ClickException("Error loading Word.Application")

@click.command()
@click.argument('input_path', type=click.Path(exists=True))
def main(input_path):
    doc = None
    try:
        print('Loading from HTML template:', input_path)
        doc = Word.Documents.Add(input_path, Visible=False)
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])

    # TODO: CLI option for output path
    output_path = input_path.replace('.html', '.docx')

    try:
        print('Saving document to:', output_path)
        doc.SaveAs(output_path, FileFormat=constants.wdFormatXMLDocument)
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])

    print("Closing document")
    doc.Close(constants.wdDoNotSaveChanges)

    print("Quitting Word")
    Word.Quit()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        Word.Quit()
        raise e
