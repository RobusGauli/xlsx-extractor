import os
import json
import argparse
from openpyxl import load_workbook

class WorkSheetNotAvailableError(Exception):

    def __init__(self, wb):
        self.message = 'Worksheet is not Avaialble.Available worksheets are \n {}' \
            .format('\n'.join(wb.get_sheet_names()))
        super().__init__(self.message)
        

def must_have(config):
    '''wrapper to verify the config json file'''
    def _wrapper(key):
        if key not in config:
            raise ValueError('{} is not provided in config file'.format(key))
        else:
            return True
    return _wrapper
        
class ExcelExtractor:

    @classmethod
    def create_from_cli(cls):
        config = cls._load_args_from_cli()
        self = cls(config)
        
        return self

    def __init__(self, config):
        self.config = config
        # get the instance of workbook file
        absolute_path = self.config.get('EXCEL_CONFIG_FILE')
        if not absolute_path or not os.path.exists(absolute_path):
            raise ValueError('Please pass down the valid excel file.')
        # now parse the json file
        self._json_config = json.loads(open(absolute_path).read())
        
        self.requires = must_have(self._json_config)
        
        self._wb = self.requires('filename') and \
            load_workbook(os.path.expanduser(self._json_config['filename']), data_only=True)

        worksheet_name = self.requires('worksheet') and self._json_config['worksheet']
        if worksheet_name not in self._wb.get_sheet_names():
            raise WorkSheetNotAvailableError(self._wb)
        self._worksheet = self._wb[worksheet_name]
        
            
    @classmethod
    def _load_args_from_cli(cls):
        config = {}
        for key, val in vars(cls._get_args_cli()).items():
            if val:
                config['EXCEL_{}'.format(key.upper())] = val
        return config
    

    @classmethod
    def _get_args_cli(cls):
        parser = argparse.ArgumentParser(
            description='Excel Sheet Extractor'
        )

        parser.add_argument(
            '-c', '--configfile',
            action='store',
            type=str,
            required=True,
            dest='config_file',
            help='Absolute Path to config file'
        )
        parser.add_argument(
            '-o', '--outputfile',
            action='store',
            type=str,
            required=True,
            dest='output_file',
            help='Absolute Path to output file'
        )
        return parser.parse_args()


    def get_rows(self):
        #this is the generator method that yield all the values from the row for each column
        start_column = self.requires('start_column') and self._json_config['start_column']
        end_column = self.requires('end_column') and self._json_config['end_column']
        
        start_row = self.requires('start_row') and self._json_config['start_row']
        end_row = self.requires('end_row') and self._json_config['end_row']

        #column_range = set(map(chr, range(ord(start_column), ord(end_column) + 1)))
        row_range = set(range(int(start_row), int(end_row) + 1))

        data_frame = self._worksheet[start_column: end_column]
        # this now has some problem 
        #we need to convert it into list
        data_frame = [[cell for cell in w] for w in data_frame]

        while all(len(col) > 0 for col in data_frame):
            #all the lenght of col must be non zero
            _result = []
            for r in data_frame:
                _cell = r.pop(0)
                if not _cell.row in row_range:
                    continue
                _result.append((_cell.value, _cell.column))
            if _result:
                yield _result
    

    def _format_to_dict_list(self):
        #get the input format
        input_format = self.requires('input') and self._json_config['input']
        output_format = self.requires('output') and self._json_config['output']
        if 'format' in self._json_config:
            global_format = self._json_config['format']
        _result = []
        for row in self.get_rows():
            _format = {}
            _input = {}
            _output = {}
            for val, col in row:
                if self._json_config['format'] and col in global_format:
                    _format[global_format[col]] = val
                if col in input_format:
                    _input[input_format[col]] = val
                if col in output_format:
                    _output[output_format[col]] = val
            _result.append({'input': _input, 'output': _output, 'chunk': _format})
        return _result
    
    def to_json(self):
        if not self.config['EXCEL_OUTPUT_FILE']:
            raise ValueError('Invalid output file')
        file_name = self.config['EXCEL_OUTPUT_FILE']
        result = self._format_to_dict_list()
        with open(file_name, mode='w') as file:
            file.write(json.dumps(result))

                



    


if __name__ == '__main__':
    e = ExcelExtractor.create_from_cli()
    e.to_json()
    
        