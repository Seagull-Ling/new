# -*- coding: utf-8 -*-
"""
撤销/重做功能模块
实现完整的状态快照管理，支持undo/redo
"""
import copy
from typing import List, Dict, Any, Optional


class HistoryManager:
    """历史状态管理器"""
    
    def __init__(self, max_history: int = 50):
        """
        初始化历史管理器
        
        Args:
            max_history: 最大历史记录数，默认为50
        """
        self.max_history = max_history
        self.undo_stack: List[Dict[str, Any]] = []
        self.redo_stack: List[Dict[str, Any]] = []
    
    def save_state(self, state: Dict[str, Any], description: str = "") -> bool:
        """
        保存当前状态到撤销栈
        
        重要：必须进行深拷贝，避免引用污染
        
        Args:
            state: 当前状态字典
            description: 状态描述（可选）
            
        Returns:
            bool: 是否保存成功
        """
        try:
            # 深拷贝状态，避免引用污染
            state_copy = self._deep_copy_state(state)
            
            # 添加描述信息
            state_copy['_description'] = description
            state_copy['_timestamp'] = self._get_timestamp()
            
            # 压入撤销栈
            self.undo_stack.append(state_copy)
            
            # 清空重做栈（新操作后不能再重做之前撤销的操作）
            self.redo_stack.clear()
            
            # 限制撤销栈大小
            if len(self.undo_stack) > self.max_history:
                self.undo_stack.pop(0)
            
            return True
            
        except Exception as e:
            print(f"保存状态失败: {e}")
            return False
    
    def undo(self) -> Optional[Dict[str, Any]]:
        """
        撤销操作，恢复到上一个状态
        
        Returns:
            Optional[Dict]: 恢复后的状态，如果无法撤销则返回None
        """
        if not self.can_undo():
            return None
        
        try:
            # 从撤销栈弹出当前状态
            current_state = self.undo_stack.pop()
            
            # 将当前状态保存到重做栈
            self.redo_stack.append(current_state)
            
            # 返回上一个状态（撤销栈的最后一个元素）
            if self.undo_stack:
                # 返回撤销栈的最新状态（深拷贝）
                return self._deep_copy_state(self.undo_stack[-1])
            else:
                # 如果撤销栈为空，返回空状态
                return None
                
        except Exception as e:
            print(f"撤销操作失败: {e}")
            return None
    
    def redo(self) -> Optional[Dict[str, Any]]:
        """
        重做操作，恢复最近一次被撤销的状态
        
        Returns:
            Optional[Dict]: 恢复后的状态，如果无法重做则返回None
        """
        if not self.can_redo():
            return None
        
        try:
            # 从重做栈弹出状态
            state_to_restore = self.redo_stack.pop()
            
            # 将状态保存回撤销栈
            self.undo_stack.append(state_to_restore)
            
            # 返回恢复的状态（深拷贝）
            return self._deep_copy_state(state_to_restore)
            
        except Exception as e:
            print(f"重做操作失败: {e}")
            return None
    
    def can_undo(self) -> bool:
        """
        检查是否可以撤销
        
        Returns:
            bool: 是否可以撤销
        """
        # 需要至少两个状态才能撤销（当前状态和上一个状态）
        # 或者说：undo_stack中有多个状态时才能撤销
        # 这里的逻辑：undo_stack中的最后一个元素是当前状态
        # 要撤销的话，需要至少有一个更早的状态
        return len(self.undo_stack) > 1
    
    def can_redo(self) -> bool:
        """
        检查是否可以重做
        
        Returns:
            bool: 是否可以重做
        """
        return len(self.redo_stack) > 0
    
    def get_undo_count(self) -> int:
        """
        获取可撤销次数
        
        Returns:
            int: 可撤销次数
        """
        return max(0, len(self.undo_stack) - 1)
    
    def get_redo_count(self) -> int:
        """
        获取可重做次数
        
        Returns:
            int: 可重做次数
        """
        return len(self.redo_stack)
    
    def get_current_state(self) -> Optional[Dict[str, Any]]:
        """
        获取当前状态（撤销栈的最后一个元素）
        
        Returns:
            Optional[Dict]: 当前状态
        """
        if self.undo_stack:
            return self._deep_copy_state(self.undo_stack[-1])
        return None
    
    def clear(self) -> None:
        """
        清空所有历史记录
        """
        self.undo_stack.clear()
        self.redo_stack.clear()
    
    def _deep_copy_state(self, state: Dict[str, Any]) -> Dict[str, Any]:
        """
        深拷贝状态字典
        
        使用copy.deepcopy进行完整深拷贝，确保没有引用共享
        
        Args:
            state: 原始状态字典
            
        Returns:
            Dict: 深拷贝后的状态字典
        """
        try:
            return copy.deepcopy(state)
        except Exception as e:
            # 如果深拷贝失败，尝试手动拷贝关键部分
            print(f"深拷贝警告: {e}，尝试手动拷贝")
            return self._manual_copy_state(state)
    
    def _manual_copy_state(self, state: Dict[str, Any]) -> Dict[str, Any]:
        """
        手动拷贝状态（当deepcopy失败时的备选方案）
        
        Args:
            state: 原始状态字典
            
        Returns:
            Dict: 拷贝后的状态字典
        """
        copied = {}
        
        # 处理基本类型
        for key, value in state.items():
            if key.startswith('_'):
                # 保留内部元数据
                copied[key] = value
            elif isinstance(value, (str, int, float, bool, type(None))):
                copied[key] = value
            elif isinstance(value, list):
                # 列表中的每个元素也需要深拷贝
                copied[key] = []
                for item in value:
                    if isinstance(item, dict):
                        copied[key].append(self._manual_copy_state(item))
                    else:
                        copied[key].append(copy.copy(item))
            elif isinstance(value, dict):
                copied[key] = self._manual_copy_state(value)
            else:
                # 其他类型尝试浅拷贝
                try:
                    copied[key] = copy.copy(value)
                except:
                    copied[key] = value
        
        return copied
    
    def _get_timestamp(self) -> str:
        """
        获取当前时间戳
        
        Returns:
            str: 时间戳字符串
        """
        from datetime import datetime
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    def get_history_descriptions(self) -> List[str]:
        """
        获取所有历史状态的描述
        
        Returns:
            List[str]: 描述列表
        """
        descriptions = []
        for state in self.undo_stack:
            desc = state.get('_description', '未命名操作')
            ts = state.get('_timestamp', '')
            if ts:
                descriptions.append(f"{desc} ({ts})")
            else:
                descriptions.append(desc)
        return descriptions
    
    def set_initial_state(self, state: Dict[str, Any]) -> None:
        """
        设置初始状态
        
        这是第一个状态，不能被撤销
        
        Args:
            state: 初始状态
        """
        self.clear()
        state_copy = self._deep_copy_state(state)
        state_copy['_description'] = '初始状态'
        state_copy['_timestamp'] = self._get_timestamp()
        self.undo_stack.append(state_copy)
