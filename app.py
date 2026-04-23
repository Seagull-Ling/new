# -*- coding: utf-8 -*-
"""
智能排座系统 - Streamlit 主应用
功能：Excel导入、姓名列选择、左侧名单编辑、右侧座位表生成、
      左右联动、手动调整、搜索、撤销/重做、导出Excel/Word
"""
import streamlit as st
from io import BytesIO
from typing import Dict, List, Optional, Any
import copy

# 导入工具模块
from utils.parser import ExcelParser
from utils.seating import SeatingGenerator
from utils.history import HistoryManager
from utils.exporter import Exporter
from utils.state_manager import StateManager


# ==================== 页面配置 ====================
st.set_page_config(
    page_title="智能排座系统",
    page_icon="🪑",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .stButton button {
        width: 100%;
        border-radius: 4px;
        transition: all 0.2s ease;
    }
    .stButton button:hover {
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }
    .seat-card {
        border: 2px solid #e0e0e0;
        border-radius: 6px;
        padding: 8px 6px;
        margin: 4px 2px;
        text-align: center;
        font-size: 12px;
        cursor: pointer;
        transition: all 0.2s ease;
        background-color: #fafafa;
    }
    .seat-card:hover {
        border-color: #4a90e2;
        background-color: #f0f7ff;
        transform: scale(1.02);
    }
    .seat-card.selected {
        border-color: #4a90e2;
        background-color: #e3f2fd;
        font-weight: bold;
    }
    .seat-card.matched {
        border-color: #ff9800;
        background-color: #fff3e0;
    }
    .seat-card.empty {
        border-style: dashed;
        color: #999;
    }
    .seat-card.swap-target {
        border-color: #f44336;
        background-color: #ffebee;
        animation: pulse 1s infinite;
    }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.7; }
    }
    .center-divider {
        background: linear-gradient(90deg, transparent, #ff6b6b, transparent);
        height: 3px;
        margin: 10px 0;
    }
    .person-row {
        padding: 8px;
        border-radius: 4px;
        margin: 4px 0;
        cursor: pointer;
        transition: all 0.2s ease;
    }
    .person-row:hover {
        background-color: #f5f5f5;
    }
    .person-row.selected {
        background-color: #e3f2fd;
        border-left: 3px solid #4a90e2;
    }
    .person-row.matched {
        background-color: #fff3e0;
        border-left: 3px solid #ff9800;
    }
    .control-section {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 15px;
        border: 1px solid #e9ecef;
    }
    .row-config-item {
        background-color: white;
        padding: 8px 12px;
        border-radius: 4px;
        margin: 4px 0;
        border: 1px solid #dee2e6;
    }
    .success-box {
        background-color: #d4edda;
        color: #155724;
        padding: 10px 15px;
        border-radius: 4px;
        border: 1px solid #c3e6cb;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        color: #856404;
        padding: 10px 15px;
        border-radius: 4px;
        border: 1px solid #ffeeba;
        margin: 10px 0;
    }
    .error-box {
        background-color: #f8d7da;
        color: #721c24;
        padding: 10px 15px;
        border-radius: 4px;
        border: 1px solid #f5c6cb;
        margin: 10px 0;
    }
    .info-box {
        background-color: #e7f3ff;
        color: #004085;
        padding: 10px 15px;
        border-radius: 4px;
        border: 1px solid #b8daff;
        margin: 10px 0;
    }
    .stats-card {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        border: 1px solid #e9ecef;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stats-number {
        font-size: 24px;
        font-weight: bold;
        color: #4a90e2;
    }
    .stats-label {
        font-size: 12px;
        color: #666;
        margin-top: 5px;
    }
</style>
""", unsafe_allow_html=True)


# ==================== 状态初始化 ====================
def init_session_state():
    """初始化 session_state"""
    if 'state_manager' not in st.session_state:
        st.session_state.state_manager = StateManager()
    
    if 'history_manager' not in st.session_state:
        st.session_state.history_manager = HistoryManager(max_history=50)
        # 保存初始状态
        initial_state = st.session_state.state_manager.get_state_snapshot()
        st.session_state.history_manager.set_initial_state(initial_state)
    
    if 'excel_parser' not in st.session_state:
        st.session_state.excel_parser = ExcelParser()
    
    if 'exporter' not in st.session_state:
        st.session_state.exporter = Exporter()
    
    if 'messages' not in st.session_state:
        st.session_state.messages = []
    
    if 'last_action' not in st.session_state:
        st.session_state.last_action = None


def save_state_to_history(description: str = ""):
    """保存当前状态到历史记录"""
    state_manager = st.session_state.state_manager
    history_manager = st.session_state.history_manager
    
    current_state = state_manager.get_state_snapshot()
    history_manager.save_state(current_state, description)
    
    st.session_state.last_action = description


def add_message(message: str, msg_type: str = "info"):
    """添加消息到消息队列"""
    st.session_state.messages.append({
        'text': message,
        'type': msg_type
    })


def show_messages():
    """显示消息"""
    for msg in st.session_state.messages:
        if msg['type'] == 'success':
            st.success(msg['text'])
        elif msg['type'] == 'warning':
            st.warning(msg['text'])
        elif msg['type'] == 'error':
            st.error(msg['text'])
        else:
            st.info(msg['text'])
    # 清空消息
    st.session_state.messages = []


# ==================== 操作函数 ====================
def handle_undo():
    """处理撤销操作"""
    history_manager = st.session_state.history_manager
    state_manager = st.session_state.state_manager
    
    if history_manager.can_undo():
        previous_state = history_manager.undo()
        if previous_state:
            state_manager.restore_from_snapshot(previous_state)
            add_message("已撤销上一步操作", "success")
    else:
        add_message("没有可撤销的操作", "warning")


def handle_redo():
    """处理重做操作"""
    history_manager = st.session_state.history_manager
    state_manager = st.session_state.state_manager
    
    if history_manager.can_redo():
        next_state = history_manager.redo()
        if next_state:
            state_manager.restore_from_snapshot(next_state)
            add_message("已重做操作", "success")
    else:
        add_message("没有可重做的操作", "warning")


def handle_reset():
    """处理重置操作"""
    state_manager = st.session_state.state_manager
    history_manager = st.session_state.history_manager
    
    state_manager.clear_all()
    history_manager.clear()
    
    # 重新设置初始状态
    initial_state = state_manager.get_state_snapshot()
    history_manager.set_initial_state(initial_state)
    
    add_message("已重置所有数据", "success")


def handle_excel_upload(uploaded_file):
    """处理Excel文件上传"""
    if uploaded_file is None:
        return
    
    excel_parser = st.session_state.excel_parser
    state_manager = st.session_state.state_manager
    
    # 读取文件
    bytes_data = uploaded_file.getvalue()
    result = excel_parser.read_excel(BytesIO(bytes_data))
    
    if result['success']:
        state_manager.set_excel_info(result['columns'], result['preview'])
        add_message(f"Excel读取成功，共 {result['total_rows']} 行数据", "success")
    else:
        add_message(result.get('error', 'Excel读取失败'), "error")


def handle_select_name_column(column_name: str):
    """处理选择姓名列"""
    excel_parser = st.session_state.excel_parser
    state_manager = st.session_state.state_manager
    
    # 提取姓名
    people, stats = excel_parser.extract_names(column_name)
    
    if stats.get('error'):
        add_message(stats['error'], "error")
        return
    
    # 更新状态管理器
    state_manager.set_people(people)
    state_manager.selected_name_column = column_name
    
    # 显示统计信息
    add_message(f"成功提取 {stats['valid_names']} 个有效姓名", "success")
    
    if stats.get('warning'):
        add_message(stats['warning'], "warning")
    
    if stats.get('has_duplicates'):
        dup_names = stats.get('duplicate_names', [])
        add_message(f"存在重复姓名: {', '.join(dup_names[:5])}{'...' if len(dup_names) > 5 else ''}", "warning")
    
    # 保存到历史记录
    save_state_to_history(f"从列'{column_name}'导入{len(people)}人")


def handle_generate_seats():
    """处理生成座位表"""
    state_manager = st.session_state.state_manager
    
    if len(state_manager.row_configs) == 0:
        add_message("请先配置每排人数", "error")
        return
    
    result = state_manager.generate_seats()
    
    if result.get('warning'):
        add_message(result['warning'], "warning")
    
    add_message(f"座位表生成成功！总座位数: {result['total_seats']}, 已安排: {result['assigned_count']}人", "success")
    
    # 保存到历史记录
    save_state_to_history(f"生成座位表({result['total_seats']}个座位)")


def handle_update_name(person_id: str, new_name: str):
    """处理更新姓名"""
    state_manager = st.session_state.state_manager
    
    if state_manager.update_person_name(person_id, new_name):
        # 如果已经生成了座位表，需要重新分配以保持同步
        if state_manager.seats:
            # 重新分配座位（保持原有分配关系，只更新姓名）
            pass
        
        save_state_to_history(f"修改姓名为'{new_name}'")


def handle_delete_person(person_id: str):
    """处理删除人员"""
    state_manager = st.session_state.state_manager
    
    person = state_manager.get_person_by_id(person_id)
    if person:
        name = person['name']
        if state_manager.delete_person(person_id):
            add_message(f"已删除: {name}", "success")
            save_state_to_history(f"删除人员'{name}'")


def handle_add_person(name: str):
    """处理添加人员"""
    state_manager = st.session_state.state_manager
    
    if not name or not name.strip():
        add_message("请输入姓名", "warning")
        return
    
    person = state_manager.add_person(name.strip())
    add_message(f"已添加: {person['name']}", "success")
    save_state_to_history(f"添加人员'{person['name']}'")


def handle_move_person(person_id: str, action: str, target_index: Optional[int] = None):
    """处理移动人员"""
    state_manager = st.session_state.state_manager
    
    person = state_manager.get_person_by_id(person_id)
    if not person:
        return
    
    name = person['name']
    success = False
    
    if action == 'up':
        success = state_manager.move_person_up(person_id)
        desc = f"上移'{name}'"
    elif action == 'down':
        success = state_manager.move_person_down(person_id)
        desc = f"下移'{name}'"
    elif action == 'top':
        success = state_manager.move_person_to_top(person_id)
        desc = f"置顶'{name}'"
    elif action == 'bottom':
        success = state_manager.move_person_to_bottom(person_id)
        desc = f"置底'{name}'"
    elif action == 'move_to' and target_index is not None:
        success = state_manager.move_person(person_id, target_index - 1)  # 转换为0-based
        desc = f"移动'{name}'到第{target_index}位"
    else:
        return
    
    if success:
        # 如果已经生成了座位表，需要重新分配
        if state_manager.seats:
            state_manager.generate_seats()
        
        save_state_to_history(desc)


def handle_select_person(person_id: Optional[str]):
    """处理选中人员"""
    state_manager = st.session_state.state_manager
    state_manager.select_person(person_id)


def handle_select_swap_seat(seat_id: str):
    """处理选择交换座位"""
    state_manager = st.session_state.state_manager
    
    if not state_manager.swap_mode_enabled:
        return
    
    success, message, first_seat = state_manager.select_swap_seat(seat_id)
    
    if first_seat and not success:
        # 只是选择了第一个座位
        add_message(message, "info")
    elif success:
        # 完成交换
        add_message(message, "success")
        save_state_to_history(f"交换座位")
    elif message:
        add_message(message, "warning")


def handle_toggle_swap_mode(enable: bool):
    """处理切换交换模式"""
    state_manager = st.session_state.state_manager
    
    if enable:
        state_manager.enable_swap_mode()
        add_message("交换模式已启用，请点击两个座位进行交换", "info")
    else:
        state_manager.disable_swap_mode()
        add_message("交换模式已取消", "info")


def handle_search(keyword: str):
    """处理搜索"""
    state_manager = st.session_state.state_manager
    state_manager.set_search_keyword(keyword)
    
    if keyword:
        results = state_manager.get_search_results()
        if results:
            add_message(f"找到 {len(results)} 个匹配人员", "success")
        else:
            add_message("未找到匹配人员", "warning")


def handle_update_row_config(row_count: int, seat_counts: Dict[int, int]):
    """处理更新每排配置"""
    state_manager = st.session_state.state_manager
    
    # 更新排数
    state_manager.update_row_count(row_count)
    
    # 更新每排座位数
    for row_no, count in seat_counts.items():
        state_manager.update_seat_count_for_row(row_no, count)


def handle_export_excel():
    """处理导出Excel"""
    state_manager = st.session_state.state_manager
    exporter = st.session_state.exporter
    
    try:
        output, validation = exporter.export_excel(
            state_manager.people,
            state_manager.seats,
            state_manager.row_configs
        )
        
        # 显示校验结果
        if validation.get('errors'):
            for err in validation['errors']:
                add_message(err, "error")
        
        if validation.get('warnings'):
            for warn in validation['warnings']:
                add_message(warn, "warning")
        
        return output, validation
        
    except Exception as e:
        add_message(f"导出失败: {str(e)}", "error")
        return None, None


def handle_export_word():
    """处理导出Word"""
    state_manager = st.session_state.state_manager
    exporter = st.session_state.exporter
    
    try:
        output, validation = exporter.export_word(
            state_manager.people,
            state_manager.seats,
            state_manager.row_configs
        )
        
        # 显示校验结果
        if validation.get('errors'):
            for err in validation['errors']:
                add_message(err, "error")
        
        if validation.get('warnings'):
            for warn in validation['warnings']:
                add_message(warn, "warning")
        
        return output, validation
        
    except Exception as e:
        add_message(f"导出失败: {str(e)}", "error")
        return None, None


# ==================== UI 渲染函数 ====================
def render_header():
    """渲染页面头部"""
    st.markdown("""
    <h1 style='text-align: center; color: #2c3e50; margin-bottom: 5px;'>
        🪑 智能排座系统
    </h1>
    <p style='text-align: center; color: #7f8c8d; margin-bottom: 20px;'>
        会议、论坛、活动排座助手 | 支持Excel导入、自动排座、手动调整、双向联动
    </p>
    """, unsafe_allow_html=True)


def render_stats():
    """渲染统计信息卡片"""
    state_manager = st.session_state.state_manager
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class='stats-card'>
            <div class='stats-number'>{state_manager.get_total_people()}</div>
            <div class='stats-label'>总人数</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='stats-card'>
            <div class='stats-number'>{state_manager.get_total_seats()}</div>
            <div class='stats-label'>总座位数</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        assigned = state_manager.get_assigned_people_count()
        st.markdown(f"""
        <div class='stats-card'>
            <div class='stats-number' style='color: {("#28a745" if assigned == state_manager.get_total_people() else "#ffc107")}'>
                {assigned}
            </div>
            <div class='stats-label'>已安排</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        rows = state_manager.get_row_count()
        st.markdown(f"""
        <div class='stats-card'>
            <div class='stats-number' style='color: #6f42c1;'>{rows}</div>
            <div class='stats-label'>总排数</div>
        </div>
        """, unsafe_allow_html=True)


def render_control_section():
    """渲染顶部控制区"""
    st.markdown("---")
    
    state_manager = st.session_state.state_manager
    history_manager = st.session_state.history_manager
    
    # 第一行：文件操作和撤销/重做
    col1, col2, col3, col4, col5, col6 = st.columns([2, 2, 1, 1, 1, 1])
    
    with col1:
        # Excel上传
        uploaded_file = st.file_uploader(
            "上传Excel文件",
            type=['xlsx'],
            key='excel_uploader'
        )
        if uploaded_file is not None:
            handle_excel_upload(uploaded_file)
    
    with col2:
        # 列选择
        if state_manager.excel_columns:
            selected_col = st.selectbox(
                "选择姓名列",
                options=['请选择...'] + state_manager.excel_columns,
                key='name_column_selector'
            )
            if selected_col != '请选择...' and selected_col != state_manager.selected_name_column:
                handle_select_name_column(selected_col)
    
    with col3:
        # 撤销按钮
        can_undo = history_manager.can_undo()
        undo_btn = st.button(
            f"↩ 撤销",
            disabled=not can_undo,
            key='undo_btn',
            use_container_width=True
        )
        if undo_btn:
            handle_undo()
    
    with col4:
        # 重做按钮
        can_redo = history_manager.can_redo()
        redo_btn = st.button(
            f"↪ 重做",
            disabled=not can_redo,
            key='redo_btn',
            use_container_width=True
        )
        if redo_btn:
            handle_redo()
    
    with col5:
        # 重置按钮
        reset_btn = st.button(
            "🔄 重置",
            key='reset_btn',
            use_container_width=True
        )
        if reset_btn:
            handle_reset()
    
    with col6:
        # 搜索框
        search_keyword = st.text_input(
            "搜索姓名",
            value=state_manager.search_keyword,
            key='search_input',
            placeholder="搜索姓名..."
        )
        if search_keyword != state_manager.search_keyword:
            handle_search(search_keyword)
    
    # 第二行：每排配置和生成按钮
    st.markdown("### 座位配置")
    config_col1, config_col2, config_col3 = st.columns([1, 3, 1])
    
    with config_col1:
        # 总排数
        current_rows = state_manager.get_row_count()
        row_count = st.number_input(
            "总排数",
            min_value=1,
            max_value=20,
            value=current_rows,
            key='row_count_input'
        )
    
    with config_col2:
        # 每排座位数配置
        st.markdown("**每排人数配置**")
        
        # 创建动态输入框
        seat_counts = {}
        cols_per_row = 5  # 每行显示多少个配置
        
        config_rows = []
        for i in range(0, row_count, cols_per_row):
            config_rows.append(state_manager.row_configs[i:i+cols_per_row])
        
        for row_idx, config_row in enumerate(config_rows):
            cols = st.columns(cols_per_row)
            for col_idx, config in enumerate(config_row):
                with cols[col_idx]:
                    seat_count = st.number_input(
                        f"第{config['row_no']}排",
                        min_value=0,
                        max_value=50,
                        value=config['seat_count'],
                        key=f"row_{config['row_no']}_seats"
                    )
                    seat_counts[config['row_no']] = seat_count
        
        # 如果排数有变化，更新配置
        if row_count != current_rows:
            handle_update_row_config(row_count, seat_counts)
        else:
            # 检查是否有座位数变化
            for row_no, new_count in seat_counts.items():
                for config in state_manager.row_configs:
                    if config['row_no'] == row_no and config['seat_count'] != new_count:
                        handle_update_row_config(row_count, seat_counts)
                        break
    
    with config_col3:
        # 生成座位表按钮
        st.markdown("<br>", unsafe_allow_html=True)
        generate_btn = st.button(
            "✨ 生成座位表",
            type='primary',
            key='generate_btn',
            use_container_width=True
        )
        if generate_btn:
            if len(state_manager.people) == 0:
                add_message("请先导入或添加人员", "warning")
            else:
                handle_generate_seats()
    
    # 第三行：导出按钮
    st.markdown("### 导出功能")
    export_col1, export_col2, export_col3 = st.columns([1, 1, 2])
    
    with export_col1:
        # 导出Excel按钮
        excel_export_btn = st.button(
            "📊 导出Excel",
            key='excel_export_btn',
            use_container_width=True
        )
        
        if excel_export_btn:
            if len(state_manager.seats) == 0:
                add_message("请先生成座位表", "warning")
            else:
                output, validation = handle_export_excel()
                if output:
                    st.download_button(
                        label="⬇ 下载Excel文件",
                        data=output,
                        file_name="座位表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key='download_excel_btn',
                        use_container_width=True
                    )
    
    with export_col2:
        # 导出Word按钮
        word_export_btn = st.button(
            "📄 导出Word",
            key='word_export_btn',
            use_container_width=True
        )
        
        if word_export_btn:
            if len(state_manager.seats) == 0:
                add_message("请先生成座位表", "warning")
            else:
                output, validation = handle_export_word()
                if output:
                    st.download_button(
                        label="⬇ 下载Word文件",
                        data=output,
                        file_name="座位表.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key='download_word_btn',
                        use_container_width=True
                    )


def render_left_panel():
    """渲染左侧面板：姓名顺序名单"""
    state_manager = st.session_state.state_manager
    
    st.markdown("### 📋 姓名顺序名单")
    
    # 添加新人员区域
    with st.expander("➕ 添加新人员", expanded=False):
        new_name_col1, new_name_col2 = st.columns([3, 1])
        with new_name_col1:
            new_name = st.text_input("输入姓名", key='new_name_input', placeholder="请输入姓名")
        with new_name_col2:
            st.markdown("<br>", unsafe_allow_html=True)
            add_btn = st.button("添加", key='add_person_btn', use_container_width=True)
            if add_btn:
                handle_add_person(new_name)
    
    # 人员列表
    people = state_manager.get_sorted_people()
    
    if not people:
        st.info("暂无人员，请上传Excel或手动添加")
        return
    
    # 遍历显示每个人
    for person in people:
        person_id = person['person_id']
        name = person['name']
        order_idx = person['order_index']
        seat_id = person.get('current_seat_id')
        
        # 判断状态
        is_selected = state_manager.is_person_selected(person_id)
        is_matched = state_manager.is_person_matched(person_id)
        
        # 获取座位信息
        seat_info = ""
        if seat_id:
            seat = state_manager.get_seat_by_id(seat_id)
            if seat:
                seat_info = f" → {seat['display_label']}"
        
        # CSS类
        css_classes = []
        if is_selected:
            css_classes.append('selected')
        if is_matched:
            css_classes.append('matched')
        class_str = ' '.join(css_classes)
        
        # 显示行
        with st.container():
            st.markdown(f"""
            <div class='person-row {class_str}' id='person-{person_id}'>
                <strong>{order_idx + 1}.</strong> {name}{seat_info}
            </div>
            """, unsafe_allow_html=True)
            
            # 操作按钮
            btn_col1, btn_col2, btn_col3, btn_col4, btn_col5, btn_col6, btn_col7 = st.columns([1, 1, 1, 1, 1, 2, 1])
            
            with btn_col1:
                if st.button("↑", key=f"up_{person_id}", use_container_width=True):
                    handle_move_person(person_id, 'up')
            
            with btn_col2:
                if st.button("↓", key=f"down_{person_id}", use_container_width=True):
                    handle_move_person(person_id, 'down')
            
            with btn_col3:
                if st.button("🔝", key=f"top_{person_id}", use_container_width=True):
                    handle_move_person(person_id, 'top')
            
            with btn_col4:
                if st.button("🔚", key=f"bottom_{person_id}", use_container_width=True):
                    handle_move_person(person_id, 'bottom')
            
            with btn_col5:
                # 编辑姓名
                with st.expander("✏️ 编辑", expanded=False):
                    edit_name = st.text_input(
                        "编辑姓名",
                        value=name,
                        key=f"edit_name_{person_id}"
                    )
                    if edit_name != name and edit_name.strip():
                        handle_update_name(person_id, edit_name)
            
            with btn_col6:
                # 移动到指定位置
                with st.expander("📍 移到", expanded=False):
                    target_pos = st.number_input(
                        "目标位置",
                        min_value=1,
                        max_value=len(people),
                        value=order_idx + 1,
                        key=f"move_pos_{person_id}"
                    )
                    if st.button("确认移动", key=f"confirm_move_{person_id}"):
                        handle_move_person(person_id, 'move_to', target_pos)
            
            with btn_col7:
                if st.button("🗑️", key=f"delete_{person_id}", use_container_width=True):
                    handle_delete_person(person_id)
            
            st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)


def render_right_panel():
    """渲染右侧面板：座位表"""
    state_manager = st.session_state.state_manager
    
    st.markdown("### 🪑 座位表")
    
    # 交换模式控制
    swap_col1, swap_col2 = st.columns([1, 3])
    
    with swap_col1:
        if state_manager.swap_mode_enabled:
            st.warning("🔄 交换模式已启用")
            if st.button("取消交换", key='cancel_swap_btn', use_container_width=True):
                handle_toggle_swap_mode(False)
        else:
            if st.button("🔄 交换座位", key='start_swap_btn', use_container_width=True):
                handle_toggle_swap_mode(True)
    
    with swap_col2:
        if state_manager.swap_mode_enabled:
            if state_manager.swap_seat1_id:
                seat1 = state_manager.get_seat_by_id(state_manager.swap_seat1_id)
                person1 = state_manager.get_person_by_seat_id(state_manager.swap_seat1_id)
                name1 = person1['name'] if person1 else '空座位'
                st.info(f"已选择: {seat1['display_label']} ({name1})，请点击第二个座位")
            else:
                st.info("请点击第一个座位")
    
    # 检查是否有座位
    if not state_manager.seats:
        st.info("请先点击\"生成座位表\"按钮")
        return
    
    # 按排显示座位
    row_numbers = state_manager.get_row_numbers()
    
    for row_no in row_numbers:
        st.markdown(f"#### 第{row_no}排")
        
        # 获取物理顺序的座位
        physical_seats = state_manager.get_row_seats_physical_order(row_no)
        
        if not physical_seats:
            st.warning(f"第{row_no}排没有座位")
            continue
        
        # 分离左右侧（用于显示中间分隔）
        left_seats = [s for s in physical_seats if s['side'] == 'LEFT']
        right_seats = [s for s in physical_seats if s['side'] == 'RIGHT']
        
        # 计算需要的列数
        total_cols = len(left_seats) + 1 + len(right_seats)  # +1是中间分隔
        
        # 创建列
        cols = st.columns(total_cols)
        
        col_idx = 0
        
        # 显示左侧座位
        for seat in left_seats:
            render_seat_card(cols[col_idx], seat, state_manager)
            col_idx += 1
        
        # 显示中间分隔
        with cols[col_idx]:
            st.markdown("""
            <div style='text-align: center; padding: 20px 5px;'>
                <div style='background: linear-gradient(90deg, transparent, #ff6b6b, transparent); 
                            height: 60px; width: 3px; margin: 0 auto; border-radius: 2px;'></div>
                <div style='font-size: 10px; color: #ff6b6b; margin-top: 5px;'>中间</div>
            </div>
            """, unsafe_allow_html=True)
        col_idx += 1
        
        # 显示右侧座位
        for seat in right_seats:
            render_seat_card(cols[col_idx], seat, state_manager)
            col_idx += 1
        
        st.markdown("---")


def render_seat_card(col, seat: Dict, state_manager: StateManager):
    """渲染单个座位卡片"""
    seat_id = seat['seat_id']
    display_label = seat['display_label']
    
    # 获取分配的人员
    person = state_manager.get_person_by_seat_id(seat_id)
    name = person['name'] if person else '（空）'
    person_id = person['person_id'] if person else None
    
    # 判断状态
    is_empty = person is None
    is_selected = state_manager.is_seat_selected(seat_id)
    is_matched = state_manager.is_seat_matched(seat_id)
    is_swap_target = (state_manager.swap_mode_enabled and 
                       state_manager.swap_seat1_id == seat_id)
    
    # 创建唯一的key
    btn_key = f"seat_{seat_id}"
    
    # 构建背景色
    bg_color = "#fafafa"
    border_color = "#e0e0e0"
    
    if is_selected:
        bg_color = "#e3f2fd"
        border_color = "#4a90e2"
    elif is_matched:
        bg_color = "#fff3e0"
        border_color = "#ff9800"
    elif is_swap_target:
        bg_color = "#ffebee"
        border_color = "#f44336"
    
    border_style = "solid"
    if is_empty:
        border_style = "dashed"
    
    text_color = "#333"
    if is_empty:
        text_color = "#999"
    
    with col:
        # 显示座位卡片（使用HTML）
        st.markdown(f"""
        <div style='text-align: center; padding: 6px 4px; border-radius: 6px; 
                    background-color: {bg_color};
                    border: 2px {border_style} {border_color};
                    margin: 4px 2px;
                    cursor: pointer;
                    transition: all 0.2s ease;'>
            <div style='font-size: 10px; color: #666; margin-bottom: 2px;'>{display_label}</div>
            <div style='font-weight: bold; font-size: 12px; color: {text_color};'>{name}</div>
        </div>
        """, unsafe_allow_html=True)
        
        # 构建按钮帮助文本
        help_text = ""
        if state_manager.swap_mode_enabled:
            if is_swap_target:
                help_text = "已选择第一个座位，请点击第二个座位完成交换"
            else:
                help_text = "点击选择此座位进行交换"
        else:
            if person_id:
                help_text = f"点击选中: {name}"
            else:
                help_text = "空座位（交换模式下可点击）"
        
        # 按钮用于处理点击（不隐藏，使用简洁文本）
        btn_label = "选中" if person_id else "点击"
        if state_manager.swap_mode_enabled:
            if is_swap_target:
                btn_label = "✓ 已选"
            else:
                btn_label = "选择"
        
        if st.button(
            btn_label,
            key=btn_key,
            help=help_text,
            use_container_width=True
        ):
            if state_manager.swap_mode_enabled:
                # 交换模式
                handle_select_swap_seat(seat_id)
            else:
                # 普通模式：选中人员
                if person_id:
                    handle_select_person(person_id)
        
        # 编辑功能（如果有人）
        if person:
            with st.expander("✏️ 编辑", expanded=False):
                edit_name = st.text_input(
                    "编辑姓名",
                    value=name,
                    key=f"edit_seat_{seat_id}"
                )
                if edit_name != name and edit_name.strip():
                    handle_update_name(person_id, edit_name)


# ==================== 主函数 ====================
def main():
    """主函数"""
    # 初始化状态
    init_session_state()
    
    # 渲染页面
    render_header()
    
    # 显示消息
    show_messages()
    
    # 统计信息
    render_stats()
    
    # 控制区
    render_control_section()
    
    st.markdown("---")
    
    # 左右分栏
    left_col, right_col = st.columns([1, 2])
    
    with left_col:
        render_left_panel()
    
    with right_col:
        render_right_panel()


if __name__ == "__main__":
    main()
