from run import default


def run_gen(model_name, out_file_path, md_file_path):
    if model_name == '默认':
        default.gen_ppt_default(out_file_path, md_file_path)
    else:
        default.gen_ppt_default(out_file_path, md_file_path)

