import os
import zipfile

def gen_annex(f_in, f_out, tid, mid, url):
    fin = open(f_in, 'rt')
    fout = open(f_out+f_in.split('/')[-1], 'wt')

    for line in fin:
        line = line.replace('my_type_id', tid)
        line = line.replace('my_mail_id', mid)
        line = line.replace('my_receive_action_url', url)
        fout.write(line)

    fin.close()
    fout.close()

def make_zip(path, zip_name):
    # ziph is zipfile handle
    # with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as ziph:
    #     for root, dirs, files in os.walk(path):
    #         for file in files:
    #             ziph.write(os.path.relpath(os.path.join(root, file), os.path.join(path, '..')))    
    
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as ziph:
        for foldername, subfolders, filenames in os.walk(path):
                if foldername == path:
                    archive_folder_name = ''
                else:
                    archive_folder_name = os.path.relpath(foldername, path)
                    ziph.write(foldername, arcname=archive_folder_name)

                for filename in filenames:
                    ziph.write(os.path.join(foldername, filename), arcname=os.path.join(archive_folder_name, filename))
        

if __name__ == '__main__':
    fin = '測資/風險金融商品比較表.doc'
    fout = '測資/附件/'

    tid, mid = '50RPribq', 'CXeyDPem'
    zipname = '測資/annex.zip'
    
    gen_annex(fin, fout, tid, mid)
    make_zip(fout, zipname)