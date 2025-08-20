# -*- coding: utf-8 -*-
"""
多語言支援模組
支援英文、繁體中文、簡體中文
"""

LANGUAGES = {
    'en': {
        'title': 'OIR Report System',
        'home': 'Home',
        'new_report': 'New Report',
        'history_report': 'History Report',
        'language': 'Language',
        'english': 'English',
        'traditional_chinese': '繁體中文',
        'simplified_chinese': '简体中文',
        
        # 新增報告
        'model_no': 'Model No.',
        'ois_no': 'OIS No.',
        'inspector': 'Inspector',
        'aaron': 'Aaron',
        'alan': 'Alan',
        'brain': 'Brain',
        'next': 'Next',
        'back': 'Back',
        'submit': 'Submit',
        'cancel': 'Cancel',
        
        # 數據輸入
        'data_input': 'Data Input',
        'item': 'Item',
        'description': 'Description',
        'datapoint': 'Datapoint',
        'enter_datapoint': 'Enter Datapoint',
        'confirm_data': 'Confirm Data',
        
        # 報告確認
        'order_no': 'Order No.',
        'shipment_size': 'Shipment Size',
        'lot_no': 'Lot No.',
        'location': 'Location',
        'date': 'Date',
        'preview': 'Preview',
        'generate_report': 'Generate Report',
        'download_excel': 'Download Excel',
        'download_pdf': 'Download PDF',
        
        # 歷史查詢
        'history_search': 'History Search',
        'date_from': 'Date From',
        'date_to': 'Date To',
        'operator': 'Operator',
        'search': 'Search',
        'no_results': 'No results found',
        'select_record': 'Select Record',
        
        # 錯誤訊息
        'error': 'Error',
        'success': 'Success',
        'ois_not_found': 'OIS No. not found in database',
        'invalid_input': 'Invalid input',
        'file_error': 'File operation error',
        'report_generated': 'Report generated successfully',
        
        # 其他
        'required': 'Required',
        'optional': 'Optional',
        'total': 'Total',
        'average': 'Average',
        'minimum': 'Minimum',
        'maximum': 'Maximum'
    },
    
    'zh-TW': {
        'title': 'OIR報告系統',
        'home': '首頁',
        'new_report': '新增報告',
        'history_report': '歷史報告',
        'language': '語言',
        'english': 'English',
        'traditional_chinese': '繁體中文',
        'simplified_chinese': '简体中文',
        
        # 新增報告
        'model_no': '型號',
        'ois_no': 'OIS編號',
        'inspector': '檢驗員',
        'aaron': 'Aaron',
        'alan': 'Alan',
        'brain': 'Brain',
        'next': '下一步',
        'back': '返回',
        'submit': '提交',
        'cancel': '取消',
        
        # 數據輸入
        'data_input': '數據輸入',
        'item': '項目',
        'description': '描述',
        'datapoint': '數據點',
        'enter_datapoint': '輸入數據點',
        'confirm_data': '確認數據',
        
        # 報告確認
        'order_no': '訂單號',
        'shipment_size': '出貨數量',
        'lot_no': '批號',
        'location': '位置',
        'date': '日期',
        'preview': '預覽',
        'generate_report': '生成報告',
        'download_excel': '下載Excel',
        'download_pdf': '下載PDF',
        
        # 歷史查詢
        'history_search': '歷史查詢',
        'date_from': '起始日期',
        'date_to': '結束日期',
        'operator': '操作員',
        'search': '搜尋',
        'no_results': '未找到結果',
        'select_record': '選擇記錄',
        
        # 錯誤訊息
        'error': '錯誤',
        'success': '成功',
        'ois_not_found': '資料庫中未找到OIS編號',
        'invalid_input': '無效輸入',
        'file_error': '檔案操作錯誤',
        'report_generated': '報告生成成功',
        
        # 其他
        'required': '必填',
        'optional': '選填',
        'total': '總計',
        'average': '平均',
        'minimum': '最小值',
        'maximum': '最大值'
    },
    
    'zh-CN': {
        'title': 'OIR报告系统',
        'home': '首页',
        'new_report': '新增报告',
        'history_report': '历史报告',
        'language': '语言',
        'english': 'English',
        'traditional_chinese': '繁體中文',
        'simplified_chinese': '简体中文',
        
        # 新增报告
        'model_no': '型号',
        'ois_no': 'OIS编号',
        'inspector': '检验员',
        'aaron': 'Aaron',
        'alan': 'Alan',
        'brain': 'Brain',
        'next': '下一步',
        'back': '返回',
        'submit': '提交',
        'cancel': '取消',
        
        # 数据输入
        'data_input': '数据输入',
        'item': '项目',
        'description': '描述',
        'datapoint': '数据点',
        'enter_datapoint': '输入数据点',
        'confirm_data': '确认数据',
        
        # 报告确认
        'order_no': '订单号',
        'shipment_size': '出货数量',
        'lot_no': '批号',
        'location': '位置',
        'date': '日期',
        'preview': '预览',
        'generate_report': '生成报告',
        'download_excel': '下载Excel',
        'download_pdf': '下载PDF',
        
        # 历史查询
        'history_search': '历史查询',
        'date_from': '起始日期',
        'date_to': '结束日期',
        'operator': '操作员',
        'search': '搜索',
        'no_results': '未找到结果',
        'select_record': '选择记录',
        
        # 错误信息
        'error': '错误',
        'success': '成功',
        'ois_not_found': '数据库中未找到OIS编号',
        'invalid_input': '无效输入',
        'file_error': '文件操作错误',
        'report_generated': '报告生成成功',
        
        # 其他
        'required': '必填',
        'optional': '选填',
        'total': '总计',
        'average': '平均',
        'minimum': '最小值',
        'maximum': '最大值'
    }
}

def get_text(key, lang='en'):
    """
    獲取指定語言的文字
    
    Args:
        key (str): 文字鍵值
        lang (str): 語言代碼 ('en', 'zh-TW', 'zh-CN')
    
    Returns:
        str: 對應語言的文字，如果找不到則返回英文版本
    """
    if lang in LANGUAGES and key in LANGUAGES[lang]:
        return LANGUAGES[lang][key]
    elif key in LANGUAGES['en']:
        return LANGUAGES['en'][key]
    else:
        return key

def get_available_languages():
    """
    獲取可用的語言列表
    
    Returns:
        list: 語言代碼列表
    """
    return list(LANGUAGES.keys())

def get_language_name(lang_code, display_lang='en'):
    """
    獲取語言的顯示名稱
    
    Args:
        lang_code (str): 語言代碼
        display_lang (str): 顯示語言
    
    Returns:
        str: 語言顯示名稱
    """
    lang_names = {
        'en': get_text('english', display_lang),
        'zh-TW': get_text('traditional_chinese', display_lang),
        'zh-CN': get_text('simplified_chinese', display_lang)
    }
    
    return lang_names.get(lang_code, lang_code)

