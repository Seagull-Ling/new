# -*- coding: utf-8 -*-
"""
页面状态初始化与同步逻辑模块
负责管理整个应用的状态，包括人员、座位、配置等
"""
from typing import List, Dict, Any, Optional, Tuple
import copy
import uuid

from .seating import SeatingGenerator


class StateManager:
    """状态管理器"""
    
    def __init__(self):
        """初始化状态管理器"""
        # 人员列表
        self.people: List[Dict] = []
        
        # 座位列表
        self.seats: List[Dict] = []
        
        # 每排配置
        self.row_configs: List[Dict] = [
            {'row_no': 1, 'seat_count': 16},
            {'row_no': 2, 'seat_count': 16},
            {'row_no': 3, 'seat_count': 16}
        ]
        
        # 搜索关键词
        self.search_keyword: str = ''
        
        # 当前选中的人员ID
        self.selected_person_id: Optional[str] = None
        
        # 交换模式相关
        self.swap_mode_enabled: bool = False
        self.swap_seat1_id: Optional[str] = None
        
        # Excel导入相关
        self.excel_columns: List[str] = []
        self.selected_name_column: Optional[str] = None
        self.excel_preview: List[Dict] = []
        
        # 排座生成器
        self.seating_generator = SeatingGenerator()
    
    def get_state_snapshot(self) -> Dict[str, Any]:
        """
        获取当前状态的完整快照（用于撤销/重做）
        
        Returns:
            Dict: 完整的状态快照
        """
        return {
            'people': copy.deepcopy(self.people),
            'seats': copy.deepcopy(self.seats),
            'row_configs': copy.deepcopy(self.row_configs),
            'search_keyword': self.search_keyword,
            'selected_person_id': self.selected_person_id,
            'swap_mode_enabled': self.swap_mode_enabled,
            'swap_seat1_id': self.swap_seat1_id,
            'excel_columns': copy.deepcopy(self.excel_columns),
            'selected_name_column': self.selected_name_column,
            'excel_preview': copy.deepcopy(self.excel_preview)
        }
    
    def restore_from_snapshot(self, snapshot: Dict[str, Any]) -> bool:
        """
        从状态快照恢复
        
        Args:
            snapshot: 状态快照
            
        Returns:
            bool: 是否恢复成功
        """
        try:
            self.people = copy.deepcopy(snapshot.get('people', []))
            self.seats = copy.deepcopy(snapshot.get('seats', []))
            self.row_configs = copy.deepcopy(snapshot.get('row_configs', []))
            self.search_keyword = snapshot.get('search_keyword', '')
            self.selected_person_id = snapshot.get('selected_person_id')
            self.swap_mode_enabled = snapshot.get('swap_mode_enabled', False)
            self.swap_seat1_id = snapshot.get('swap_seat1_id')
            self.excel_columns = copy.deepcopy(snapshot.get('excel_columns', []))
            self.selected_name_column = snapshot.get('selected_name_column')
            self.excel_preview = copy.deepcopy(snapshot.get('excel_preview', []))
            
            # 重新初始化排座生成器
            self.seating_generator = SeatingGenerator()
            if self.seats:
                # 重新生成座位结构
                self.seating_generator.generate_seat_structure(self.row_configs)
                # 恢复座位分配
                for seat in self.seating_generator.seats:
                    for saved_seat in self.seats:
                        if seat['seat_id'] == saved_seat['seat_id']:
                            seat['assigned_person_id'] = saved_seat.get('assigned_person_id')
                            break
                self.seats = self.seating_generator.seats
            
            return True
            
        except Exception as e:
            print(f"恢复状态失败: {e}")
            return False
    
    def set_people(self, people: List[Dict]) -> None:
        """
        设置人员列表
        
        Args:
            people: 人员列表
        """
        self.people = copy.deepcopy(people)
    
    def add_person(self, name: str, position: Optional[int] = None) -> Dict:
        """
        添加新人员
        
        Args:
            name: 姓名
            position: 插入位置（None表示添加到末尾）
            
        Returns:
            Dict: 新增的人员信息
        """
        person = {
            'person_id': str(uuid.uuid4()),
            'name': name.strip(),
            'order_index': len(self.people) if position is None else position,
            'current_seat_id': None
        }
        
        if position is None:
            self.people.append(person)
        else:
            # 插入到指定位置，并更新后续人员的order_index
            self.people.insert(position, person)
            # 重新计算所有order_index
            for idx, p in enumerate(self.people):
                p['order_index'] = idx
        
        return person
    
    def delete_person(self, person_id: str) -> bool:
        """
        删除人员
        
        Args:
            person_id: 人员ID
            
        Returns:
            bool: 是否删除成功
        """
        for i, person in enumerate(self.people):
            if person['person_id'] == person_id:
                # 从座位中移除
                for seat in self.seats:
                    if seat.get('assigned_person_id') == person_id:
                        seat['assigned_person_id'] = None
                
                # 删除人员
                self.people.pop(i)
                
                # 重新计算order_index
                for idx, p in enumerate(self.people):
                    p['order_index'] = idx
                
                # 如果删除的是当前选中的人员，清除选中状态
                if self.selected_person_id == person_id:
                    self.selected_person_id = None
                
                return True
        
        return False
    
    def update_person_name(self, person_id: str, new_name: str) -> bool:
        """
        更新人员姓名
        
        Args:
            person_id: 人员ID
            new_name: 新姓名
            
        Returns:
            bool: 是否更新成功
        """
        for person in self.people:
            if person['person_id'] == person_id:
                person['name'] = new_name.strip()
                return True
        return False
    
    def move_person(self, person_id: str, target_index: int) -> bool:
        """
        移动人员到指定位置
        
        Args:
            person_id: 人员ID
            target_index: 目标位置（从0开始）
            
        Returns:
            bool: 是否移动成功
        """
        # 找到当前位置
        current_index = -1
        person_to_move = None
        
        for i, person in enumerate(self.people):
            if person['person_id'] == person_id:
                current_index = i
                person_to_move = person
                break
        
        if current_index == -1 or person_to_move is None:
            return False
        
        # 边界检查
        if target_index < 0:
            target_index = 0
        if target_index >= len(self.people):
            target_index = len(self.people) - 1
        
        # 如果位置没有变化
        if current_index == target_index:
            return True
        
        # 移动人员
        self.people.pop(current_index)
        self.people.insert(target_index, person_to_move)
        
        # 重新计算order_index
        for idx, p in enumerate(self.people):
            p['order_index'] = idx
        
        return True
    
    def move_person_up(self, person_id: str) -> bool:
        """
        上移人员
        
        Args:
            person_id: 人员ID
            
        Returns:
            bool: 是否移动成功
        """
        for i, person in enumerate(self.people):
            if person['person_id'] == person_id:
                if i > 0:
                    return self.move_person(person_id, i - 1)
                return False
        return False
    
    def move_person_down(self, person_id: str) -> bool:
        """
        下移人员
        
        Args:
            person_id: 人员ID
            
        Returns:
            bool: 是否移动成功
        """
        for i, person in enumerate(self.people):
            if person['person_id'] == person_id:
                if i < len(self.people) - 1:
                    return self.move_person(person_id, i + 1)
                return False
        return False
    
    def move_person_to_top(self, person_id: str) -> bool:
        """
        人员置顶
        
        Args:
            person_id: 人员ID
            
        Returns:
            bool: 是否移动成功
        """
        return self.move_person(person_id, 0)
    
    def move_person_to_bottom(self, person_id: str) -> bool:
        """
        人员置底
        
        Args:
            person_id: 人员ID
            
        Returns:
            bool: 是否移动成功
        """
        return self.move_person(person_id, len(self.people) - 1)
    
    def generate_seats(self) -> Dict[str, Any]:
        """
        根据当前配置生成座位并分配人员
        
        Returns:
            Dict: 生成结果信息
        """
        # 生成座位结构
        self.seating_generator.generate_seat_structure(self.row_configs)
        
        # 分配人员到座位
        self.seats, self.people = self.seating_generator.assign_people_to_seats(self.people)
        
        total_seats = len(self.seats)
        total_people = len(self.people)
        
        result = {
            'success': True,
            'total_seats': total_seats,
            'total_people': total_people,
            'assigned_count': min(total_people, total_seats),
            'unassigned_count': max(0, total_people - total_seats)
        }
        
        if total_people > total_seats:
            result['warning'] = f"人数({total_people})超过座位数({total_seats})，有{total_people - total_seats}人未安排座位"
        
        return result
    
    def set_row_configs(self, configs: List[Dict]) -> None:
        """
        设置每排配置
        
        Args:
            configs: 每排配置列表
        """
        self.row_configs = copy.deepcopy(configs)
    
    def update_row_count(self, new_count: int) -> None:
        """
        更新总排数
        
        Args:
            new_count: 新的排数
        """
        current_count = len(self.row_configs)
        
        if new_count > current_count:
            # 增加排数
            for i in range(current_count, new_count):
                self.row_configs.append({
                    'row_no': i + 1,
                    'seat_count': 16  # 默认16个座位
                })
        elif new_count < current_count:
            # 减少排数
            self.row_configs = self.row_configs[:new_count]
    
    def update_seat_count_for_row(self, row_no: int, seat_count: int) -> bool:
        """
        更新指定排的座位数
        
        Args:
            row_no: 排号
            seat_count: 新的座位数
            
        Returns:
            bool: 是否更新成功
        """
        for config in self.row_configs:
            if config['row_no'] == row_no:
                config['seat_count'] = max(0, seat_count)
                return True
        return False
    
    def get_row_seats_physical_order(self, row_no: int) -> List[Dict]:
        """
        获取某排座位的物理显示顺序
        
        Args:
            row_no: 排号
            
        Returns:
            List[Dict]: 按物理顺序排列的座位列表
        """
        return self.seating_generator.get_physical_order_seats(row_no)
    
    def get_seat_by_id(self, seat_id: str) -> Optional[Dict]:
        """
        根据seat_id获取座位
        """
        return self.seating_generator.get_seat_by_id(seat_id)
    
    def get_person_by_id(self, person_id: str) -> Optional[Dict]:
        """
        根据person_id获取人员
        """
        for person in self.people:
            if person['person_id'] == person_id:
                return person
        return None
    
    def get_person_by_seat_id(self, seat_id: str) -> Optional[Dict]:
        """
        根据座位ID获取分配的人员
        """
        seat = self.get_seat_by_id(seat_id)
        if not seat:
            return None
        
        person_id = seat.get('assigned_person_id')
        if not person_id:
            return None
        
        return self.get_person_by_id(person_id)
    
    def swap_seats(self, seat_id1: str, seat_id2: str) -> Tuple[bool, str]:
        """
        交换两个座位的人员
        
        这是"右改左跟"的核心操作
        
        Args:
            seat_id1: 第一个座位ID
            seat_id2: 第二个座位ID
            
        Returns:
            Tuple[bool, str]: (是否成功, 消息)
        """
        if seat_id1 == seat_id2:
            return False, "不能交换同一个座位"
        
        seat1 = self.get_seat_by_id(seat_id1)
        seat2 = self.get_seat_by_id(seat_id2)
        
        if not seat1 or not seat2:
            return False, "座位不存在"
        
        # 执行交换
        self.seating_generator.swap_seats(seat_id1, seat_id2, self.people)
        
        # 根据座位分配重新计算人员顺序（右改左跟）
        self.people = self.seating_generator.reorder_people_by_seats(self.people)
        
        return True, "交换成功"
    
    def enable_swap_mode(self) -> None:
        """启用交换模式"""
        self.swap_mode_enabled = True
        self.swap_seat1_id = None
    
    def disable_swap_mode(self) -> None:
        """禁用交换模式"""
        self.swap_mode_enabled = False
        self.swap_seat1_id = None
    
    def select_swap_seat(self, seat_id: str) -> Tuple[bool, str, Optional[str]]:
        """
        选择交换座位
        
        Args:
            seat_id: 座位ID
            
        Returns:
            Tuple[bool, str, Optional[str]]: (是否完成交换, 消息, 第一个座位ID)
        """
        if not self.swap_mode_enabled:
            return False, "交换模式未启用", None
        
        if self.swap_seat1_id is None:
            # 选择第一个座位
            self.swap_seat1_id = seat_id
            seat = self.get_seat_by_id(seat_id)
            person = self.get_person_by_seat_id(seat_id)
            name = person['name'] if person else '空座位'
            return False, f"已选择第一个座位: {seat['display_label']} ({name})，请选择第二个座位", seat_id
        else:
            # 选择第二个座位，执行交换
            if self.swap_seat1_id == seat_id:
                self.swap_seat1_id = None
                return False, "不能选择同一个座位，请重新选择", None
            
            # 执行交换
            success, message = self.swap_seats(self.swap_seat1_id, seat_id)
            
            # 重置交换状态
            first_seat_id = self.swap_seat1_id
            self.swap_seat1_id = None
            self.swap_mode_enabled = False
            
            return success, message, first_seat_id
    
    def set_search_keyword(self, keyword: str) -> None:
        """
        设置搜索关键词
        
        Args:
            keyword: 搜索关键词
        """
        self.search_keyword = keyword.strip()
    
    def get_search_results(self) -> List[Dict]:
        """
        获取搜索结果
        
        Returns:
            List[Dict]: 匹配的人员列表
        """
        if not self.search_keyword:
            return []
        
        keyword = self.search_keyword.lower()
        results = []
        
        for person in self.people:
            if keyword in person['name'].lower():
                results.append(person)
        
        return results
    
    def is_person_matched(self, person_id: str) -> bool:
        """
        检查人员是否匹配搜索结果
        
        Args:
            person_id: 人员ID
            
        Returns:
            bool: 是否匹配
        """
        if not self.search_keyword:
            return False
        
        person = self.get_person_by_id(person_id)
        if not person:
            return False
        
        return self.search_keyword.lower() in person['name'].lower()
    
    def is_seat_matched(self, seat_id: str) -> bool:
        """
        检查座位上的人员是否匹配搜索结果
        
        Args:
            seat_id: 座位ID
            
        Returns:
            bool: 是否匹配
        """
        person = self.get_person_by_seat_id(seat_id)
        if not person:
            return False
        
        return self.is_person_matched(person['person_id'])
    
    def select_person(self, person_id: Optional[str]) -> None:
        """
        选中/取消选中人员
        
        Args:
            person_id: 人员ID（None表示取消选中）
        """
        self.selected_person_id = person_id
    
    def is_person_selected(self, person_id: str) -> bool:
        """
        检查人员是否被选中
        
        Args:
            person_id: 人员ID
            
        Returns:
            bool: 是否被选中
        """
        return self.selected_person_id == person_id
    
    def is_seat_selected(self, seat_id: str) -> bool:
        """
        检查座位上的人员是否被选中
        
        Args:
            seat_id: 座位ID
            
        Returns:
            bool: 是否被选中
        """
        person = self.get_person_by_seat_id(seat_id)
        if not person:
            return False
        
        return self.is_person_selected(person['person_id'])
    
    def clear_all(self) -> None:
        """
        清空所有数据（重置）
        """
        self.people = []
        self.seats = []
        self.row_configs = [
            {'row_no': 1, 'seat_count': 16},
            {'row_no': 2, 'seat_count': 16},
            {'row_no': 3, 'seat_count': 16}
        ]
        self.search_keyword = ''
        self.selected_person_id = None
        self.swap_mode_enabled = False
        self.swap_seat1_id = None
        self.excel_columns = []
        self.selected_name_column = None
        self.excel_preview = []
        self.seating_generator = SeatingGenerator()
    
    def get_sorted_people(self) -> List[Dict]:
        """
        获取按顺序排列的人员列表
        
        Returns:
            List[Dict]: 排序后的人员列表
        """
        return sorted(self.people, key=lambda x: x['order_index'])
    
    def get_total_seats(self) -> int:
        """
        获取总座位数
        """
        return len(self.seats)
    
    def get_total_people(self) -> int:
        """
        获取总人数
        """
        return len(self.people)
    
    def get_assigned_people_count(self) -> int:
        """
        获取已分配座位的人数
        """
        return sum(1 for p in self.people if p.get('current_seat_id'))
    
    def get_row_count(self) -> int:
        """
        获取总排数
        """
        return len(self.row_configs)
    
    def get_row_numbers(self) -> List[int]:
        """
        获取所有排号列表
        """
        return [config['row_no'] for config in self.row_configs]
    
    def set_excel_info(self, columns: List[str], preview: List[Dict]) -> None:
        """
        设置Excel导入相关信息
        
        Args:
            columns: 列名列表
            preview: 预览数据
        """
        self.excel_columns = copy.deepcopy(columns)
        self.excel_preview = copy.deepcopy(preview)
