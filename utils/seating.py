# -*- coding: utf-8 -*-
"""
核心排座算法
实现从中间向两侧、先右后左的排座规则
"""
from typing import List, Dict, Tuple, Optional
import copy


class SeatingGenerator:
    """座位生成器"""
    
    def __init__(self):
        self.seats: List[Dict] = []
        self.row_configs: List[Dict] = []
    
    def generate_seat_structure(self, row_configs: List[Dict]) -> List[Dict]:
        """
        根据每排配置生成座位结构
        
        Args:
            row_configs: 每排配置列表，每个元素包含 'row_no' 和 'seat_count'
            
        Returns:
            List[Dict]: 座位列表，按排座顺序排列
        """
        self.row_configs = copy.deepcopy(row_configs)
        self.seats = []
        
        for config in row_configs:
            row_no = config['row_no']
            seat_count = config['seat_count']
            
            if seat_count <= 0:
                continue
            
            # 生成该排的座位
            row_seats = self._generate_row_seats(row_no, seat_count)
            self.seats.extend(row_seats)
        
        return self.seats
    
    def _generate_row_seats(self, row_no: int, seat_count: int) -> List[Dict]:
        """
        生成单排座位
        
        排座规则：
        - 以中间区域为分界线
        - 每排座位分为左右两侧
        - 排座顺序固定为：右1、左2、右3、左4、右5、左6……
        - 奇数人数时，优先从中间向外安排，顺序保持"先右后左"
        
        Args:
            row_no: 排号
            seat_count: 该排座位数
            
        Returns:
            List[Dict]: 该排座位列表，按排座顺序排列
        """
        seats = []
        
        # 计算左右两侧座位数
        # 偶数：左右各一半
        # 奇数：右侧比左侧多1个（或相等，取决于具体实现）
        # 这里采用：奇数时，右侧比左侧多1个，中间的座位属于右侧
        right_count = (seat_count + 1) // 2
        left_count = seat_count // 2
        
        # 按照排座顺序生成座位：右1、左1、右2、左2、右3、左3...
        # 注意：这里的顺序是排座优先级顺序，不是物理位置顺序
        # 物理位置顺序在显示时需要调整
        
        seat_order = 0  # 排座顺序索引
        
        for i in range(max(right_count, left_count)):
            # 先右后左
            # 右侧座位（编号从1开始）
            if i < right_count:
                seat_no = i + 1
                seat = {
                    'seat_id': f'R{row_no}_RIGHT_{seat_no}',
                    'row_no': row_no,
                    'side': 'RIGHT',
                    'seat_no': seat_no,
                    'display_label': f'第{row_no}排右{seat_no}',
                    'assigned_person_id': None,
                    'seat_order': seat_order  # 排座顺序
                }
                seats.append(seat)
                seat_order += 1
            
            # 左侧座位（编号从1开始）
            if i < left_count:
                seat_no = i + 1
                seat = {
                    'seat_id': f'R{row_no}_LEFT_{seat_no}',
                    'row_no': row_no,
                    'side': 'LEFT',
                    'seat_no': seat_no,
                    'display_label': f'第{row_no}排左{seat_no}',
                    'assigned_person_id': None,
                    'seat_order': seat_order  # 排座顺序
                }
                seats.append(seat)
                seat_order += 1
        
        return seats
    
    def get_physical_order_seats(self, row_no: int) -> List[Dict]:
        """
        获取某排座位的物理显示顺序
        
        物理显示顺序：
        - 从左到右：左最大、左次大...左2、左1、【中间】、右1、右2...右最大
        - 例如：左4、左3、左2、左1、中间、右1、右2、右3、右4
        
        Args:
            row_no: 排号
            
        Returns:
            List[Dict]: 按物理顺序排列的座位列表
        """
        row_seats = [s for s in self.seats if s['row_no'] == row_no]
        
        # 分离左右两侧
        left_seats = [s for s in row_seats if s['side'] == 'LEFT']
        right_seats = [s for s in row_seats if s['side'] == 'RIGHT']
        
        # 左侧座位按seat_no从大到小排列（左4、左3、左2、左1）
        left_seats_sorted = sorted(left_seats, key=lambda x: x['seat_no'], reverse=True)
        
        # 右侧座位按seat_no从小到大排列（右1、右2、右3、右4）
        right_seats_sorted = sorted(right_seats, key=lambda x: x['seat_no'])
        
        # 合并物理顺序：左侧倒序 + 右侧正序
        # 中间分界会在UI层处理
        physical_order = left_seats_sorted + right_seats_sorted
        
        return physical_order
    
    def assign_people_to_seats(self, people: List[Dict]) -> Tuple[List[Dict], List[Dict]]:
        """
        将人员分配到座位
        
        按照排座顺序依次分配人员
        如果人员数大于座位数，多余人员不分配
        
        Args:
            people: 人员列表，按顺序排列
            
        Returns:
            Tuple[List[Dict], List[Dict]]: (更新后的座位列表, 更新后的人员列表)
        """
        # 先清空所有座位
        for seat in self.seats:
            seat['assigned_person_id'] = None
        
        # 按排座顺序分配
        # 注意：需要按seat_order排序，而不是物理顺序
        seats_by_order = sorted(self.seats, key=lambda x: x['seat_order'])
        
        # 更新人员的座位信息
        for person in people:
            person['current_seat_id'] = None
        
        # 分配人员到座位
        for i, person in enumerate(people):
            if i < len(seats_by_order):
                seat = seats_by_order[i]
                seat['assigned_person_id'] = person['person_id']
                person['current_seat_id'] = seat['seat_id']
        
        return self.seats, people
    
    def get_seat_by_id(self, seat_id: str) -> Optional[Dict]:
        """
        根据seat_id获取座位信息
        """
        for seat in self.seats:
            if seat['seat_id'] == seat_id:
                return seat
        return None
    
    def get_total_seats(self) -> int:
        """
        获取总座位数
        """
        return len(self.seats)
    
    def get_row_count(self) -> int:
        """
        获取总排数
        """
        if not self.seats:
            return 0
        return max(seat['row_no'] for seat in self.seats)
    
    def get_row_seats_count(self, row_no: int) -> int:
        """
        获取某排的座位数
        """
        return len([s for s in self.seats if s['row_no'] == row_no])
    
    def swap_seats(self, seat_id1: str, seat_id2: str, people: List[Dict]) -> Tuple[List[Dict], List[Dict]]:
        """
        交换两个座位的人员
        
        如果其中一个座位为空，则将人员从一个座位移动到另一个座位
        
        Args:
            seat_id1: 第一个座位ID
            seat_id2: 第二个座位ID
            people: 人员列表
            
        Returns:
            Tuple[List[Dict], List[Dict]]: (更新后的座位列表, 更新后的人员列表)
        """
        seat1 = self.get_seat_by_id(seat_id1)
        seat2 = self.get_seat_by_id(seat_id2)
        
        if not seat1 or not seat2:
            return self.seats, people
        
        # 保存当前分配
        person_id1 = seat1['assigned_person_id']
        person_id2 = seat2['assigned_person_id']
        
        # 交换分配
        seat1['assigned_person_id'] = person_id2
        seat2['assigned_person_id'] = person_id1
        
        # 更新人员的座位信息
        for person in people:
            if person['person_id'] == person_id1:
                person['current_seat_id'] = seat_id2
            elif person['person_id'] == person_id2:
                person['current_seat_id'] = seat_id1
        
        return self.seats, people
    
    def reorder_people_by_seats(self, people: List[Dict]) -> List[Dict]:
        """
        根据当前座位分配重新计算人员顺序
        
        这是"右改左跟"的核心逻辑：
        当右侧座位表发生变化后，左侧名单顺序应该根据座位分配重新计算
        
        规则：
        1. 已分配座位的人员按排座顺序排列
        2. 未分配座位的人员按原顺序排在后面
        
        Args:
            people: 原始人员列表
            
        Returns:
            List[Dict]: 重新排序后的人员列表
        """
        # 按排座顺序获取座位
        seats_by_order = sorted(self.seats, key=lambda x: x['seat_order'])
        
        # 创建人员ID到人员的映射
        person_map = {p['person_id']: p for p in people}
        
        # 收集已分配座位的人员（按排座顺序）
        assigned_people = []
        assigned_person_ids = set()
        
        for seat in seats_by_order:
            person_id = seat.get('assigned_person_id')
            if person_id and person_id in person_map:
                person = person_map[person_id]
                assigned_people.append(person)
                assigned_person_ids.add(person_id)
        
        # 收集未分配座位的人员（保持原顺序）
        unassigned_people = [p for p in people if p['person_id'] not in assigned_person_ids]
        
        # 合并并更新order_index
        result = assigned_people + unassigned_people
        
        for idx, person in enumerate(result):
            person['order_index'] = idx
        
        return result
    
    def get_person_by_id(self, person_id: str, people: List[Dict]) -> Optional[Dict]:
        """
        根据person_id获取人员信息
        """
        for person in people:
            if person['person_id'] == person_id:
                return person
        return None
