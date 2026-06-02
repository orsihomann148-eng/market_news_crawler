#!/usr/bin/env python3
from __future__ import annotations

import sys
from pathlib import Path
from typing import Callable

from country_config import DEFAULT_COUNTRY_CODE, country_options, get_country_config, normalize_country_code
import credential_manager
import news_crawler
import source_manager
import xlsx_source_test

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / 'outputs'

PLATFORM_CHOICES = [
    ('1', 'tiktok_shop', 'TikTok Shop'),
    ('2', 'amazon_japan', 'Amazon Japan'),
    ('3', 'rakuten_ichiba', 'Rakuten Ichiba'),
    ('4', 'qoo10', 'Qoo10'),
    ('5', 'temu', 'TEMU'),
    ('6', 'shein', 'SHEIN'),
]

SIDE_CHOICES = [
    ('1', 'media', '媒体侧'),
    ('2', 'buyer', '买家侧'),
    ('3', 'seller', '卖家侧'),
]

COUNTRY_CHOICES = [
    (str(index), code, label)
    for index, (code, label) in enumerate(country_options(), start=1)
]


def line(char: str = '=') -> None:
    print(char * 72)


def title(text: str) -> None:
    line()
    print(text)
    line()


def ask(prompt: str, default: str = '') -> str:
    suffix = f' [{default}]' if default else ''
    value = input(f'{prompt}{suffix}: ').strip()
    return value or default


def ask_int(prompt: str, default: int) -> int:
    raw = ask(prompt, str(default))
    try:
        return int(raw)
    except ValueError:
        print(f'输入无效，已使用默认值 {default}。')
        return default


def ask_yes_no(prompt: str, default: bool = True) -> bool:
    hint = 'Y/n' if default else 'y/N'
    raw = input(f'{prompt} [{hint}]: ').strip().lower()
    if not raw:
        return default
    return raw in {'y', 'yes', '1', 'true'}


def choose_many(options: list[tuple[str, str, str]], prompt_text: str) -> list[str]:
    print(prompt_text)
    for no, value, label in options:
        print(f'  {no}. {label} ({value})')
    raw = input('请输入编号，多个用英文逗号分隔；直接回车表示全选: ').strip()
    if not raw:
        return [value for _, value, _ in options]
    selected: list[str] = []
    mapping = {no: value for no, value, _ in options}
    for part in raw.split(','):
        key = part.strip()
        if key in mapping and mapping[key] not in selected:
            selected.append(mapping[key])
    if not selected:
        print('未识别到有效编号，已默认全选。')
        return [value for _, value, _ in options]
    return selected


def choose_country(default: str = DEFAULT_COUNTRY_CODE) -> str:
    print('选择国家：')
    for no, value, label in COUNTRY_CHOICES:
        print(f'  {no}. {label} ({value})')
    raw = input(f'请输入编号，直接回车使用默认值 {default}: ').strip()
    if not raw:
        return default
    mapping = {no: value for no, value, _ in COUNTRY_CHOICES}
    return normalize_country_code(mapping.get(raw, raw))


def run_with_result(label: str, func: Callable[[list[str] | None], int], argv: list[str]) -> None:
    line('-')
    print(f'正在执行: {label}')
    print('命令参数:', ' '.join(argv) if argv else '(无)')
    line('-')
    try:
        code = func(argv)
    except SystemExit as exc:
        code = int(exc.code) if isinstance(exc.code, int) else 1
    print()
    print(f'执行完成，退出码: {code}')
    print()


def show_latest_outputs() -> None:
    title('最近输出结果')
    if not OUTPUT_DIR.exists():
        print('当前还没有 outputs 目录。')
        return
    folders = [path for path in OUTPUT_DIR.iterdir() if path.is_dir()]
    folders.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    if not folders:
        print('当前还没有运行输出。')
        return
    for folder in folders[:10]:
        print(folder.name)
        for child in sorted(folder.iterdir()):
            print(f'  - {child.name}')
        print()


def run_news_crawler_wizard() -> None:
    title('官方新闻抓取向导（日本专用）')
    mode = ask('选择时间模式：1=近 N 天，2=指定起止日期', '1')
    argv: list[str] = []
    if mode == '2':
        start_date = ask('开始日期（YYYY-MM-DD）')
        end_date = ask('结束日期（YYYY-MM-DD）')
        if start_date:
            argv += ['--start-date', start_date]
        if end_date:
            argv += ['--end-date', end_date]
    else:
        days = ask_int('抓取最近多少天', 7)
        argv += ['--days', str(days)]

    platforms = choose_many(PLATFORM_CHOICES, '选择要抓取的平台：')
    if len(platforms) < len(PLATFORM_CHOICES):
        for platform in platforms:
            argv += ['--platform', platform]

    translate_to = ask('翻译语言（zh-CN 或 en）', 'zh-CN')
    output_dir = ask('输出目录', 'outputs')
    if translate_to:
        argv += ['--translate-to', translate_to]
    if output_dir:
        argv += ['--output-dir', output_dir]

    run_with_result('news_crawler.py', news_crawler.main, argv)


def run_xlsx_wizard() -> None:
    title('表格来源抓取向导')
    country_code = choose_country()
    country_config = get_country_config(country_code)
    argv: list[str] = []
    argv += ['--country', country_code]
    argv += ['--xlsx', ask('表格路径（默认在国家子目录）', str(country_config['xlsx_path']))]
    argv += ['--days', str(ask_int('抓取最近多少天', 7))]
    sides = choose_many(SIDE_CHOICES, '选择要抓取的侧别：')
    if sides:
        argv += ['--sides', *sides]
    argv += ['--translate-to', ask('翻译语言（zh-CN 或 en）', 'zh-CN')]
    argv += ['--max-links-per-source', str(ask_int('每个来源最多跟进多少个候选链接', 8))]
    argv += ['--workers', str(ask_int('并发线程数', 8))]
    argv += ['--extra-sources', ask('额外来源配置文件（默认在国家子目录）', str(country_config['extra_sources_path']))]
    argv += ['--adapter-configs', ask('站点适配配置文件（默认在国家子目录）', str(country_config['adapter_configs_path']))]
    argv += ['--site-credentials', ask('站点凭据配置文件（默认在国家子目录）', str(country_config['site_credentials_path']))]
    argv += ['--output-dir', ask('输出目录', 'outputs')]
    run_with_result('xlsx_source_test.py', xlsx_source_test.main, argv)


def run_source_manager_wizard() -> None:
    title('来源管理向导')
    country_code = choose_country()
    country_config = get_country_config(country_code)
    mode = ask('选择操作：1=新增来源，2=停用来源，3=查看列表', '1')
    argv: list[str] = [
        '--extra-sources', str(country_config['extra_sources_path']),
        '--adapter-configs', str(country_config['adapter_configs_path']),
        '--capability-cache', str(country_config['source_capability_cache_path']),
    ]
    if mode == '3':
        argv.append('--list')
        side = ask('筛选侧别（media/buyer/seller，留空表示全部）', '')
        platform = ask('筛选平台名（留空表示全部）', '')
        domain = ask('筛选域名（留空表示全部）', '')
        if side:
            argv += ['--side', side]
        if platform:
            argv += ['--platform', platform]
        if domain:
            argv += ['--domain', domain]
        if ask_yes_no('是否显示已停用记录', False):
            argv.append('--show-inactive')
    elif mode == '2':
        argv.append('--remove')
        url = ask('要停用的网址（可留空，仅按平台/侧别/域名匹配）', '')
        platform = ask('平台名（可留空）', '')
        side = ask('侧别 media/buyer/seller（可留空）', '')
        domain = ask('域名（可留空）', '')
        if url:
            argv.append(url)
        if platform:
            argv += ['--platform', platform]
        if side:
            argv += ['--side', side]
        if domain:
            argv += ['--domain', domain]
    else:
        url = ask('新增的网址')
        platform = ask('平台名，例如 TikTok/TikTok Shop')
        side = ask('侧别 media/buyer/seller', 'media')
        argv += [url, '--platform', platform, '--side', side]
        argv += ['--extra-sources', ask('额外来源配置文件（默认在国家子目录）', str(country_config['extra_sources_path']))]
        argv += ['--adapter-configs', ask('站点适配配置文件（默认在国家子目录）', str(country_config['adapter_configs_path']))]
        if ask_yes_no('是否跳过 API 适配器生成', False):
            argv.append('--skip-api')
        else:
            api_url = ask('API URL（留空则读取环境变量）', '')
            api_key = ask('API Key（留空则读取环境变量）', '')
            api_model = ask('API Model（留空则读取环境变量）', '')
            if api_url:
                argv += ['--api-url', api_url]
            if api_key:
                argv += ['--api-key', api_key]
            if api_model:
                argv += ['--api-model', api_model]
            if ask_yes_no('如果域名已有适配配置，是否强制重新生成', False):
                argv.append('--force-api')
    run_with_result('source_manager.py', source_manager.main, argv)


def run_credential_manager_wizard() -> None:
    title('站点凭据管理向导')
    country_code = choose_country()
    country_config = get_country_config(country_code)
    mode = ask('选择操作：1=查看凭据，2=新增/更新凭据，3=清空凭据', '1')
    argv: list[str] = ['--credentials-file', str(country_config['site_credentials_path'])]
    if mode == '1':
        argv.append('--list')
    else:
        target = ask('输入域名或网址，例如 sellercentral.amazon.co.jp')
        argv.append(target)
        if mode == '3':
            argv.append('--clear')
        else:
            auth_type = ask('认证方式 form/basic/cookie', 'cookie')
            argv += ['--auth-type', auth_type]
            username = ask('用户名（可留空）', '')
            password = ask('密码（可留空）', '')
            cookie_header = ask('Cookie Header（可留空）', '')
            header = ask('自定义请求头，格式 Name: Value（可留空）', '')
            note = ask('备注（可留空）', '')
            if username:
                argv += ['--username', username]
            if password:
                argv += ['--password', password]
            if cookie_header:
                argv += ['--cookie-header', cookie_header]
            if header:
                argv += ['--header', header]
            if note:
                argv += ['--note', note]
            if ask_yes_no('是否启用该凭据', True):
                argv.append('--enable')
            else:
                argv.append('--disable')
    run_with_result('credential_manager.py', credential_manager.main, argv)


def show_quick_start() -> None:
    title('快速开始')
    print('更友好的使用方式：')
    print('1. 直接运行 python3 app.py')
    print('2. 按菜单编号选择功能')
    print('3. 根据提示输入日期、平台、路径或凭据')
    print('4. 程序会自动拼接原命令行参数并执行')
    print()
    print('仍然支持原来的命令行方式，例如：')
    print('  python3 news_crawler.py --days 14 --platform qoo10')
    print('  python3 xlsx_source_test.py --days 30 --sides media buyer')
    print('  python3 xlsx_source_test.py --country france --days 30 --sides media buyer')
    print('  python3 source_manager.py --list --show-inactive')
    print('  python3 credential_manager.py --list')
    print()


def main() -> int:
    while True:
        title('多国家新闻资讯抓取工具')
        print('1. 抓取官方新闻（news_crawler，日本专用）')
        print('2. 抓取表格来源（xlsx_source_test）')
        print('3. 管理来源网址（source_manager）')
        print('4. 管理登录凭据（credential_manager）')
        print('5. 查看最近输出结果')
        print('6. 查看快速开始说明')
        print('0. 退出')
        choice = input('请选择功能编号: ').strip()
        print()
        if choice == '1':
            run_news_crawler_wizard()
        elif choice == '2':
            run_xlsx_wizard()
        elif choice == '3':
            run_source_manager_wizard()
        elif choice == '4':
            run_credential_manager_wizard()
        elif choice == '5':
            show_latest_outputs()
        elif choice == '6':
            show_quick_start()
        elif choice == '0':
            print('已退出。')
            return 0
        else:
            print('请输入有效编号。')
        input('按回车键继续...')
        print()


if __name__ == '__main__':
    sys.exit(main())
