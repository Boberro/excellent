from pyramid.view import view_config
from StringIO import StringIO
import pyramid_excel as pyexcel
import pyexcel.ext.xls
import pyexcel.ext.xlsx
import pyexcel.ext.ods


def includeme(config):
    config.add_route('home', '/')
    config.add_route('csv', '/{filename}.csv')
    config.add_route('xls', '/{filename}.xls')
    config.add_route('xlsx', '/{filename}.xlsx')
    config.add_route('ods', '/{filename}.ods')


@view_config(route_name='home', renderer='templates/index.jinja2')
def my_view(request):
    return {'project': 'excellent'}


def generate_sheet(file_format):
    io = StringIO()

    sheet_data = [
        ["Kolumna 1", "Kolumna 2", "Kolumna Zygmunta"],
        [1, 2, "Zupa"],
        ['I', 'III', 'Dupa'],
    ]
    sheet = pyexcel.Sheet(sheet_data)
    sheet.save_to_memory(file_format, io)
    return io.getvalue()


@view_config(route_name='csv')
def csv_view(request):
    filename = request.matchdict['filename']

    request.response.body = generate_sheet('csv')
    request.response.content_type = 'text/csv'
    request.response.content_disposition = 'attachment; filename={0}.csv'.format(filename)

    return request.response


@view_config(route_name='xls')
def xls_view(request):
    filename = request.matchdict['filename']

    request.response.body = generate_sheet('xls')
    request.response.content_type = 'application/vnd.ms-excel'
    request.response.content_disposition = 'attachment; filename={0}.xls'.format(filename)

    return request.response


@view_config(route_name='xlsx')
def xlsx_view(request):
    filename = request.matchdict['filename']

    request.response.body = generate_sheet('xlsx')
    request.response.content_type = 'application/vnd.ms-excel'
    request.response.content_disposition = 'attachment; filename={0}.xlsx'.format(filename)

    return request.response


@view_config(route_name='ods')
def odt_view(request):
    filename = request.matchdict['filename']

    request.response.body = generate_sheet('ods')
    request.response.content_type = 'application/vnd.oasis.opendocument.spreadsheet'
    request.response.content_disposition = 'attachment; filename={0}.ods'.format(filename)

    return request.response
