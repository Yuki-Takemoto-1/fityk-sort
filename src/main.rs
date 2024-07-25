use fityk_sort_rust::{io::read_toml_file, sort_peaks};

const DEFAULT_SETTINGS_FILENAME: &str = "fityk_option.txt";
fn main() {
    let args: Vec<String> = std::env::args().collect();
    let (_, commands) = args.split_first().unwrap();
    let filename = match commands.first() {
        Some(filename) => filename,
        None => {
            println!("設定ファイルがコマンドラインで指定されなかったため,デフォルトのファイル({DEFAULT_SETTINGS_FILENAME})を読み込みます");
            DEFAULT_SETTINGS_FILENAME
        }
    };

    let toml_data = read_toml_file(filename).unwrap();

    if let Err(err) = &toml_data.options.create_dir() {
        println!("フォルダ作成時にエラーが発生しました:{err:?}");
        return;
    }

    for filename in toml_data.settings.filenames.iter() {
        if let Err(err) = sort_peaks(filename, &toml_data.settings.columns, &toml_data.options) {
            println!("エラーが発生しました:{err:?}")
        } else {
            println!("{filename:?}の処理が完了しました")
        }
    }
}
