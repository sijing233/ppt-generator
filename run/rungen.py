from run import default
from pptx.dml.color import RGBColor

import os


model_dict = {
    "通用": {
        'model_name': 'default-model',
        'other_rgb': RGBColor(208, 206, 206),
        'current_rgb': RGBColor(66, 85, 108),
        'content_number_rgb': RGBColor(255, 255, 255),
        'second_title_rgb': RGBColor(66, 85, 108)
    }
}


def run_gen(model_name, out_file_path, md_file_path):
    project_path = '..'
    # project_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if model_name == '通用':
        default.gen_ppt_default(model_dict[model_name]['model_name'], out_file_path, md_file_path, project_path,
                                model_dict[model_name]['other_rgb'],
                                model_dict[model_name]['current_rgb'],
                                model_dict[model_name]['content_number_rgb'],
                                model_dict[model_name]['second_title_rgb'])
    else:
        model_name = '通用'
        default.gen_ppt_default(model_dict[model_name]['model_name'], out_file_path, md_file_path, project_path,
                                model_dict[model_name]['other_rgb'],
                                model_dict[model_name]['current_rgb'],
                                model_dict[model_name]['content_number_rgb'],
                                model_dict[model_name]['second_title_rgb'])
