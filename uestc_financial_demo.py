#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电子科技大学财务系统自动化演示
基于实际的财务综合信息门户
"""

import asyncio
import json
import os
import requests
from datetime import datetime
from typing import Dict, List, Optional
from dataclasses import dataclass
import logging

# Playwright相关导入
from playwright.async_api import async_playwright, Browser, Page

# 导入配置
import config
import pandas as pd
CAPTCHA_MODULE = "manual"  # 手动输入验证码

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@dataclass
class UserInputData:
    """用户输入数据类"""
    name: str
    project_number: str
    amount: float

@dataclass
class ExpenseItem:
    """报销项目数据类"""
    project: str
    account: str
    amount: float
    description: str
    date: str
    category: str
    vendor: str = ""
    invoice_number: str = ""

class UESTCFinancialAutomation:
    """电子科技大学财务系统自动化"""
    
    def __init__(self):
        self.expenses: List[ExpenseItem] = []
        self.browser: Optional[Browser] = None
        self.page: Optional[Page] = None
        self.config = config.FINANCIAL_SYSTEM_CONFIG
        self.project_config = config.PROJECT_CONFIG
        self.is_logged_in = False
        # 验证码处理方式：手动输入
        
    def read_excel_expense_data(self) -> dict:
        """读取Excel文件中的报销业务数据"""
        try:
            # 读取Excel文件的Sheet_Baoxiao sheet，指定编码
            df = pd.read_excel('报销信息.xlsx', sheet_name='Sheet_Baoxiao', engine='openpyxl')
            
            if len(df) == 0:
                logger.warning("Sheet_Baoxiao sheet为空")
                return {}
            
            # 获取第一行数据（跳过表头）
            first_row = df.iloc[0]  # 这里已经是第一行数据，因为pandas会自动跳过表头
            
            expense_data = {}
            
            # 读取项目编号
            if '项目编号' in df.columns:
                expense_data['project_number'] = str(first_row['项目编号'])
                logger.info(f"✓ 从Excel文件读取项目编号: {expense_data['project_number']}")
            
            # 读取附件张数
            if '附件张数' in df.columns:
                expense_data['attachment_count'] = int(first_row['附件张数'])
                logger.info(f"✓ 从Excel文件读取附件张数: {expense_data['attachment_count']}")
            
            # 读取支付方式
            if '支付方式' in df.columns:
                payment_method = first_row['支付方式']
                logger.info(f"原始支付方式值: {payment_method}, 类型: {type(payment_method)}")
                
                if pd.notna(payment_method) and str(payment_method).strip() != '':  # 检查是否为NaN或空字符串
                    expense_data['payment_method'] = str(payment_method).strip()
                    logger.info(f"✓ 从Excel文件读取支付方式: {expense_data['payment_method']}")
                else:
                    # 如果Excel中支付方式为空，使用默认值
                    expense_data['payment_method'] = "个人转卡"  # 默认支付方式
                    logger.info(f"✓ 使用默认支付方式: {expense_data['payment_method']}")
            
            # 读取金额
            if '金额' in df.columns:
                expense_data['amount'] = float(first_row['金额'])
                logger.info(f"✓ 从Excel文件读取金额: {expense_data['amount']}")
            
            # 读取预约科目
            if '预约科目' in df.columns:
                appointment_subject = first_row['预约科目']
                logger.info(f"原始预约科目值: {appointment_subject}, 类型: {type(appointment_subject)}")
                
                if pd.notna(appointment_subject) and str(appointment_subject).strip() != '':
                    expense_data['appointment_subject'] = str(appointment_subject).strip()
                    logger.info(f"✓ 从Excel文件读取预约科目: {expense_data['appointment_subject']}")
                else:
                    # 如果Excel中预约科目为空，使用默认值
                    expense_data['appointment_subject'] = "差旅费"
                    logger.info(f"✓ 使用默认预约科目: {expense_data['appointment_subject']}")
            else:
                # 如果没有预约科目列，使用默认值
                expense_data['appointment_subject'] = "差旅费"
                logger.info(f"✓ 使用默认预约科目: {expense_data['appointment_subject']}")
            
            # 读取工号
            if '工号' in df.columns:
                employee_id = first_row['工号']
                if pd.notna(employee_id) and str(employee_id).strip() != '':
                    expense_data['employee_id'] = str(employee_id).strip()
                    logger.info(f"✓ 从Excel文件读取工号: {expense_data['employee_id']}")
                else:
                    expense_data['employee_id'] = ""
                    logger.info("✓ 工号为空，使用空字符串")
            else:
                expense_data['employee_id'] = ""
                logger.info("✓ 未找到工号列，使用空字符串")
            
            # 读取个人
            if '个人' in df.columns:
                personal_name = first_row['个人']
                if pd.notna(personal_name) and str(personal_name).strip() != '':
                    expense_data['personal_name'] = str(personal_name).strip()
                    logger.info(f"✓ 从Excel文件读取个人: {expense_data['personal_name']}")
                else:
                    expense_data['personal_name'] = ""
                    logger.info("✓ 个人为空，使用空字符串")
            else:
                expense_data['personal_name'] = ""
                logger.info("✓ 未找到个人列，使用空字符串")
            
            # 读取卡号
            if '卡号' in df.columns:
                card_number = first_row['卡号']
                if pd.notna(card_number) and str(card_number).strip() != '':
                    expense_data['card_number'] = str(card_number).strip()
                    logger.info(f"✓ 从Excel文件读取卡号: {expense_data['card_number']}")
                else:
                    expense_data['card_number'] = ""
                    logger.info("✓ 卡号为空，使用空字符串")
            else:
                expense_data['card_number'] = ""
                logger.info("✓ 未找到卡号列，使用空字符串")
            
            # 读取个人金额
            if '个人金额' in df.columns:
                personal_amount = first_row['个人金额']
                if pd.notna(personal_amount) and str(personal_amount).strip() != '':
                    try:
                        expense_data['personal_amount'] = float(personal_amount)
                        logger.info(f"✓ 从Excel文件读取个人金额: {expense_data['personal_amount']}")
                    except (ValueError, TypeError):
                        expense_data['personal_amount'] = 0.0
                        logger.info("✓ 个人金额格式错误，使用0.0")
                else:
                    expense_data['personal_amount'] = 0.0
                    logger.info("✓ 个人金额为空，使用0.0")
            else:
                expense_data['personal_amount'] = 0.0
                logger.info("✓ 未找到个人金额列，使用0.0")
            
            return expense_data
                
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
            return {}
    
    async def fill_expense_form(self) -> bool:
        """填写报销表单"""
        try:
            logger.info("=== 填写报销表单 ===")
            print("=== 开始填写报销表单 ===")
            
            # 等待页面完全加载
            await asyncio.sleep(2)
            
            # 调试：显示当前页面的输入框信息
            logger.info("调试：分析当前页面输入框...")
            print("调试：分析当前页面输入框...")
            await self.debug_input_fields()
            
            # 读取Excel数据
            logger.info("开始读取Excel数据...")
            print("开始读取Excel数据...")
            expense_data = self.read_excel_expense_data()
            if not expense_data:
                logger.error("无法读取Excel数据")
                print("✗ 无法读取Excel数据")
                return False
            
            logger.info(f"从Excel读取的数据: {expense_data}")
            print(f"✓ 从Excel读取的数据: {expense_data}")
            
            # 填写报销项目号
            logger.info("开始填写项目编号...")
            print("开始填写项目编号...")
            project_result = await self.fill_project_number(expense_data["project_number"])
            if project_result:
                print("✓ 项目编号填写成功")
            else:
                logger.warning("填写报销项目号失败，但继续执行后续步骤")
                print("⚠ 填写报销项目号失败，但继续执行后续步骤")
            
            # 填写附件张数
            logger.info("开始填写附件张数...")
            print("开始填写附件张数...")
            attachment_result = await self.fill_attachment_count(expense_data["attachment_count"])
            if attachment_result:
                print("✓ 附件张数填写成功")
            else:
                logger.warning("填写附件张数失败，但继续执行后续步骤")
                print("⚠ 填写附件张数失败，但继续执行后续步骤")
            
            # 选择支付方式
            logger.info("开始选择支付方式...")
            print("开始选择支付方式...")
            payment_result = await self.select_payment_method(expense_data["payment_method"])
            if payment_result:
                print("✓ 支付方式选择成功")
            else:
                logger.warning("选择支付方式失败，但继续执行后续步骤")
                print("⚠ 选择支付方式失败，但继续执行后续步骤")
            
            logger.info("✓ 基础表单填写完成")
            print("✓ 基础表单填写完成")
            
            # 点击下一步按钮
            logger.info("点击下一步按钮...")
            print("点击下一步按钮...")
            next_result = await self.click_next_button()
            if next_result:
                print("✓ 下一步按钮点击成功")
            else:
                print("⚠ 下一步按钮点击失败")
            
            # 等待页面加载（增加到10秒）
            logger.info("等待页面加载（10秒）...")
            print("等待页面加载（10秒）...")
            await asyncio.sleep(10)
            
            # 使用LLM智能填写预约科目金额
            logger.info("开始LLM智能填写预约科目...")
            print("开始LLM智能填写预约科目...")
            llm_result = await self.fill_appointment_subjects_with_llm(expense_data)
            if llm_result:
                print("✓ LLM智能填写预约科目成功")
            else:
                logger.warning("智能填写预约科目失败")
                print("⚠ 智能填写预约科目失败")
            
            logger.info("✓ 报销表单填写流程完成")
            print("✓ 报销表单填写流程完成")
            return True
                
        except Exception as e:
            logger.error(f"填写报销表单失败: {e}")
            return False

    def read_subject_mapping(self) -> dict:
        """读取科目-输入框ID对应表"""
        try:
            import pandas as pd
            logger.info("正在读取科目-输入框ID对应表...")
            
            # 读取Excel文件
            df = pd.read_excel('科目-输入框id对应.xlsx', engine='openpyxl')
            
            logger.info(f"Excel文件列名: {list(df.columns)}")
            logger.info(f"数据行数: {len(df)}")
            
            # 创建科目到输入框ID和含义说明的映射字典
            subject_mapping = {}
            
            # 使用正确的列名
            subject_name_col = '科目名称（b_name）'
            input_id_col = '输入框ID（value输入框）'
            description_col = None
            
            # 查找科目含义说明列
            for col in df.columns:
                if '说明' in col or '含义' in col or '描述' in col or '备注' in col:
                    description_col = col
                    logger.info(f"找到科目含义说明列: {description_col}")
                    break
            
            if subject_name_col in df.columns and input_id_col in df.columns:
                for i, row in df.iterrows():
                    subject_name = str(row[subject_name_col]).strip()
                    input_id = str(row[input_id_col]).strip()
                    
                    # 读取科目含义说明
                    description = ""
                    if description_col and description_col in df.columns:
                        description = str(row[description_col]).strip()
                        if description == 'nan':
                            description = ""
                    
                    if subject_name and input_id and subject_name != 'nan' and input_id != 'nan':
                        subject_mapping[subject_name] = {
                            'input_id': input_id,
                            'description': description
                        }
                        logger.info(f"映射: {subject_name} -> {input_id} (说明: {description})")
            
            logger.info(f"总共创建了 {len(subject_mapping)} 个映射关系")
            return subject_mapping
            
        except Exception as e:
            logger.error(f"读取科目映射表失败: {e}")
            return {}

    async def fill_appointment_subjects_with_llm(self, expense_data: dict) -> bool:
        """使用LLM智能填写预约科目金额"""
        try:
            logger.info("=== 使用LLM智能填写预约科目金额 ===")
            print("=== 使用LLM智能填写预约科目金额 ===")
            
            # 读取科目-输入框ID对应表
            print("正在读取科目-输入框ID对应表...")
            subject_mapping = self.read_subject_mapping()
            if not subject_mapping:
                logger.error("无法读取科目映射表")
                print("✗ 无法读取科目映射表")
                return False
            
            logger.info(f"科目映射表包含 {len(subject_mapping)} 个映射关系")
            print(f"✓ 科目映射表包含 {len(subject_mapping)} 个映射关系")
            
            # 获取页面上的预约科目信息
            print("正在获取页面上的预约科目信息...")
            subjects_info = await self.get_appointment_subjects_info()
            if not subjects_info:
                logger.error("无法获取预约科目信息")
                print("✗ 无法获取预约科目信息")
                return False
            
            logger.info(f"获取到的预约科目信息: {subjects_info}")
            print(f"✓ 获取到的预约科目信息: {len(subjects_info)} 个科目")
            
            # 使用LLM分析并确定最合适的科目
            print("正在使用LLM分析并确定最合适的科目...")
            target_subject = await self.analyze_with_llm(
                expense_data["appointment_subject"], 
                expense_data["amount"], 
                subjects_info,
                subject_mapping
            )
            
            if not target_subject:
                logger.error("LLM分析失败")
                print("✗ LLM分析失败")
                return False
            
            logger.info(f"LLM推荐的目标科目: {target_subject}")
            print(f"✓ LLM推荐的目标科目: {target_subject['name']}")
            
            # 填写金额到目标科目
            print("正在填写金额到目标科目...")
            if not await self.fill_amount_to_subject(target_subject, expense_data["amount"]):
                logger.error("填写金额到科目失败")
                print("✗ 填写金额到科目失败")
                return False
            
            # 输出LLM匹配结果
            print(f"\n🎯 LLM匹配结果: 报销单中的预约科目 '{expense_data['appointment_subject']}' 已通过LLM智能匹配到界面中的预约科目 '{target_subject['name']}'，并填写金额 ¥{expense_data['amount']} 到输入框 {target_subject['input_selector']}")
            
            logger.info("✓ 预约科目金额填写完成")
            print("✓ 预约科目金额填写完成")
            
            # 等待一下确保金额填写完成
            await asyncio.sleep(2)
            
            # 点击下一步按钮
            logger.info("准备点击下一步按钮...")
            print("准备点击下一步按钮...")
            if await self.click_next_button():
                logger.info("✓ 成功点击下一步按钮，进入下一个界面")
                print(f"\n✅ 已成功点击下一步按钮，进入下一个界面")
                
                # 等待页面加载
                await asyncio.sleep(3)
                
                # 填写个人信息表单（工号、个人、卡号、个人金额）
                logger.info("开始填写个人信息表单...")
                print("开始填写个人信息表单...")
                if await self.fill_personal_info_form(expense_data):
                    logger.info("✓ 个人信息表单填写完成")
                    print(f"\n✅ 个人信息表单填写完成")
                else:
                    logger.warning("个人信息表单填写失败")
                    print(f"\n⚠️ 个人信息表单填写失败")
                
                return True
            else:
                logger.warning("点击下一步按钮失败")
                print(f"\n⚠️ 点击下一步按钮失败，但金额填写已完成")
                return False
            
        except Exception as e:
            logger.error(f"智能填写预约科目失败: {e}")
            return False

    async def fill_personal_info_form(self, expense_data: dict) -> bool:
        """填写个人信息表单（工号、个人、卡号、个人金额）"""
        try:
            logger.info("=== 填写个人信息表单 ===")
            
            # 等待页面加载
            await asyncio.sleep(3)
            
            # 调试：显示当前页面的输入框信息
            logger.info("调试：分析当前页面输入框...")
            await self.debug_input_fields()
            
            success_count = 0
            total_fields = 0
            
            # 填写工号
            if expense_data.get('employee_id'):
                total_fields += 1
                logger.info(f"开始填写工号: {expense_data['employee_id']}")
                if await self.fill_employee_id(expense_data['employee_id']):
                    success_count += 1
                    logger.info("✓ 工号填写成功")
                else:
                    logger.warning("✗ 工号填写失败")
            
            # 填写个人姓名
            if expense_data.get('personal_name'):
                total_fields += 1
                logger.info(f"开始填写个人姓名: {expense_data['personal_name']}")
                if await self.fill_personal_name(expense_data['personal_name']):
                    success_count += 1
                    logger.info("✓ 个人姓名填写成功")
                else:
                    logger.warning("✗ 个人姓名填写失败")
            
            # 填写卡号
            if expense_data.get('card_number'):
                total_fields += 1
                logger.info(f"开始填写卡号: {expense_data['card_number']}")
                if await self.fill_card_number(expense_data['card_number']):
                    success_count += 1
                    logger.info("✓ 卡号填写成功")
                else:
                    logger.warning("✗ 卡号填写失败")
            
            # 填写个人金额
            if expense_data.get('personal_amount', 0) > 0:
                total_fields += 1
                logger.info(f"开始填写个人金额: {expense_data['personal_amount']}")
                if await self.fill_personal_amount(expense_data['personal_amount']):
                    success_count += 1
                    logger.info("✓ 个人金额填写成功")
                else:
                    logger.warning("✗ 个人金额填写失败")
            
            logger.info(f"个人信息填写完成: {success_count}/{total_fields} 个字段填写成功")
            
            if total_fields > 0:
                print(f"\n📝 个人信息填写结果: {success_count}/{total_fields} 个字段填写成功")
                if success_count == total_fields:
                    print("✅ 所有个人信息字段填写完成")
                    return True
                else:
                    print("⚠️ 部分个人信息字段填写失败，但继续执行")
                    return True  # 即使部分失败也继续
            else:
                logger.info("没有需要填写的个人信息字段")
                return True
                
        except Exception as e:
            logger.error(f"填写个人信息表单失败: {e}")
            return False

    async def fill_employee_id(self, employee_id: str) -> bool:
        """填写工号"""
        try:
            logger.info(f"开始填写工号: {employee_id}")
            
            # 首先尝试使用指定的ID
            specific_selector = "#formWF_YB6_3950_ypt-sno"
            
            try:
                # 等待页面加载
                await asyncio.sleep(1)
                
                # 查找工号输入框
                employee_input = self.page.locator(specific_selector)
                if await employee_input.count() > 0:
                    await employee_input.fill(employee_id)
                    logger.info(f"✓ 使用指定ID填写工号成功: {specific_selector}")
                    
                    # 等待可能出现的银行卡选择对话框
                    await asyncio.sleep(2)
                    
                    # 检查是否有银行卡选择对话框出现
                    bank_dialog = self.page.locator("#paybankdiv")
                    if await bank_dialog.count() > 0:
                        logger.info("检测到银行卡选择对话框，开始处理...")
                        return await self.handle_bank_card_selection()
                    else:
                        logger.info("未检测到银行卡选择对话框")
                        return True
                else:
                    logger.warning(f"未找到指定的工号输入框: {specific_selector}")
            except Exception as e:
                logger.warning(f"使用指定ID填写工号失败: {e}")
            
            # 如果指定ID失败，使用通用选择器
            logger.info("尝试使用通用选择器填写工号...")
            selectors = [
                "input[name*='employee']",
                "input[name*='工号']",
                "input[id*='employee']",
                "input[id*='工号']",
                "input[placeholder*='工号']",
                "input[placeholder*='employee']",
                "input[title*='工号']",
                "input[title*='employee']",
                "input[type='text']"  # 通用文本输入框
            ]
            
            return await self.fill_input_field(selectors, employee_id, "工号")
            
        except Exception as e:
            logger.error(f"填写工号失败: {e}")
            return False

    async def handle_bank_card_selection(self) -> bool:
        """处理银行卡选择对话框"""
        try:
            logger.info("=== 处理银行卡选择对话框 ===")
            
            # 读取Excel数据获取卡号信息
            expense_data = self.read_excel_expense_data()
            target_card_number = expense_data.get('card_number', '')
            
            if not target_card_number:
                logger.warning("Excel中没有卡号信息，无法进行银行卡匹配")
                return True  # 继续执行，不阻塞流程
            
            logger.info(f"目标卡号: {target_card_number}")
            
            # 等待对话框完全加载
            await asyncio.sleep(2)
            
            # 查找银行卡选择对话框
            bank_dialog = self.page.locator("#paybankdiv")
            if await bank_dialog.count() == 0:
                logger.warning("未找到银行卡选择对话框")
                return True
            
            # 查找所有银行卡选项
            bank_rows = self.page.locator("#paybankdiv table tbody tr")
            row_count = await bank_rows.count()
            
            logger.info(f"找到 {row_count} 个银行卡选项")
            
            # 跳过表头行，从第二行开始处理
            for i in range(1, row_count):
                try:
                    row = bank_rows.nth(i)
                    
                    # 获取该行的银行卡信息
                    cells = row.locator("td")
                    if await cells.count() >= 4:
                        # 获取姓名、卡号、卡类型、联行号
                        name = await cells.nth(1).text_content()
                        card_number = await cells.nth(2).text_content()
                        card_type = await cells.nth(3).text_content()
                        bank_info = await cells.nth(4).text_content()
                        
                        name = name.strip() if name else ""
                        card_number = card_number.strip() if card_number else ""
                        card_type = card_type.strip() if card_type else ""
                        bank_info = bank_info.strip() if bank_info else ""
                        
                        logger.info(f"银行卡选项 {i}: 姓名={name}, 卡号={card_number}, 类型={card_type}, 联行号={bank_info}")
                        
                        # 匹配逻辑：检查卡号是否匹配
                        if self.match_card_number(target_card_number, card_number):
                            logger.info(f"✓ 找到匹配的银行卡: {card_number}")
                            
                            # 点击对应的单选按钮
                            radio_button = row.locator("input[type='radio']")
                            if await radio_button.count() > 0:
                                await radio_button.click()
                                logger.info("✓ 成功选择银行卡")
                                print(f"\n✅ 已成功选择银行卡: {card_number}")
                                
                                # 等待选择生效
                                await asyncio.sleep(1)
                                return True
                            else:
                                logger.warning("未找到单选按钮")
                        else:
                            logger.debug(f"卡号不匹配: 目标={target_card_number}, 当前={card_number}")
                
                except Exception as e:
                    logger.warning(f"处理银行卡选项 {i} 时出错: {e}")
                    continue
            
            logger.warning("未找到匹配的银行卡")
            print(f"\n⚠️ 未找到匹配的银行卡，目标卡号: {target_card_number}")
            return True  # 即使没找到匹配的也继续执行
            
        except Exception as e:
            logger.error(f"处理银行卡选择对话框失败: {e}")
            return True  # 出错也继续执行

    def match_card_number(self, target_card: str, display_card: str) -> bool:
        """匹配银行卡号"""
        try:
            if not target_card or not display_card:
                return False
            
            # 清理卡号格式
            target_clean = target_card.replace(' ', '').replace('-', '').strip()
            display_clean = display_card.replace(' ', '').replace('-', '').strip()
            
            # 如果显示的是掩码格式（如 6227******1142），需要特殊处理
            if '*' in display_clean:
                # 提取显示卡号的前4位和后4位
                parts = display_clean.split('*')
                if len(parts) >= 2:
                    prefix = parts[0]
                    suffix = parts[-1]
                    
                    # 检查目标卡号是否匹配前4位和后4位
                    if (len(target_clean) >= 8 and 
                        target_clean.startswith(prefix) and 
                        target_clean.endswith(suffix)):
                        logger.info(f"✓ 卡号匹配成功: 目标={target_clean}, 显示={display_clean}")
                        return True
            
            # 直接比较（如果显示的是完整卡号）
            if target_clean == display_clean:
                logger.info(f"✓ 卡号完全匹配: {target_clean}")
                return True
            
            # 如果目标卡号比显示卡号长，检查是否包含显示卡号
            if len(target_clean) > len(display_clean) and target_clean.endswith(display_clean):
                logger.info(f"✓ 卡号后缀匹配: 目标={target_clean}, 显示={display_clean}")
                return True
            
            return False
            
        except Exception as e:
            logger.error(f"卡号匹配失败: {e}")
            return False

    async def fill_personal_name(self, personal_name: str) -> bool:
        """填写个人姓名"""
        try:
            # 个人姓名输入框的选择器列表
            selectors = [
                "input[name*='personal']",
                "input[name*='个人']",
                "input[name*='name']",
                "input[name*='姓名']",
                "input[id*='personal']",
                "input[id*='个人']",
                "input[id*='name']",
                "input[id*='姓名']",
                "input[placeholder*='个人']",
                "input[placeholder*='姓名']",
                "input[placeholder*='personal']",
                "input[placeholder*='name']",
                "input[title*='个人']",
                "input[title*='姓名']",
                "input[type='text']"  # 通用文本输入框
            ]
            
            return await self.fill_input_field(selectors, personal_name, "个人姓名")
            
        except Exception as e:
            logger.error(f"填写个人姓名失败: {e}")
            return False

    async def fill_card_number(self, card_number: str) -> bool:
        """填写卡号"""
        try:
            # 卡号输入框的选择器列表
            selectors = [
                "input[name*='card']",
                "input[name*='卡号']",
                "input[name*='bank']",
                "input[name*='银行']",
                "input[id*='card']",
                "input[id*='卡号']",
                "input[id*='bank']",
                "input[id*='银行']",
                "input[placeholder*='卡号']",
                "input[placeholder*='card']",
                "input[placeholder*='银行']",
                "input[placeholder*='bank']",
                "input[title*='卡号']",
                "input[title*='card']",
                "input[type='text']"  # 通用文本输入框
            ]
            
            return await self.fill_input_field(selectors, card_number, "卡号")
            
        except Exception as e:
            logger.error(f"填写卡号失败: {e}")
            return False

    async def fill_personal_amount(self, personal_amount: float) -> bool:
        """填写个人金额"""
        try:
            # 个人金额输入框的选择器列表
            selectors = [
                "input[name*='personal_amount']",
                "input[name*='个人金额']",
                "input[name*='amount']",
                "input[name*='金额']",
                "input[id*='personal_amount']",
                "input[id*='个人金额']",
                "input[id*='amount']",
                "input[id*='金额']",
                "input[placeholder*='个人金额']",
                "input[placeholder*='amount']",
                "input[placeholder*='金额']",
                "input[title*='个人金额']",
                "input[title*='amount']",
                "input[title*='金额']",
                "input[type='number']",  # 数字输入框
                "input[type='text']"     # 通用文本输入框
            ]
            
            return await self.fill_input_field(selectors, str(personal_amount), "个人金额")
            
        except Exception as e:
            logger.error(f"填写个人金额失败: {e}")
            return False

    async def get_appointment_subjects_info(self) -> list:
        """获取页面上的预约科目信息"""
        try:
            subjects = []
            
            # 首先尝试从科目映射表中获取科目信息
            subject_mapping = self.read_subject_mapping()
            if subject_mapping:
                logger.info("从科目映射表中获取科目信息...")
                for subject_name, mapping_info in subject_mapping.items():
                    input_id = mapping_info['input_id']
                    description = mapping_info.get('description', f"科目映射表中的科目: {subject_name}")
                    subjects.append({
                        "id": f"mapped_{subject_name}",
                        "name": subject_name,
                        "description": description,
                        "input_selector": f"#{input_id}"
                    })
                    logger.info(f"从映射表添加科目: {subject_name} -> {input_id} (说明: {description})")
            
            # 如果映射表中没有科目，再尝试从页面中查找
            if not subjects:
                logger.info("映射表中没有科目，尝试从页面中查找...")
                frames = self.page.frames
                for frame_idx, frame in enumerate(frames):
                    if frame.url and "WF_YB6" in frame.url:
                        logger.info(f"在iframe {frame_idx} 中查找预约科目")
                        
                        # 查找所有可编辑的科目行
                        subject_rows = await frame.locator("tr[id^='B']").all()
                        
                        for row in subject_rows:
                            try:
                                # 获取科目名称
                                name_cell = await row.locator("td[aria-describedby*='t.b_name'] span").first
                                if await name_cell.is_visible():
                                    subject_name = await name_cell.text_content()
                                    subject_name = subject_name.strip() if subject_name else ""
                                    
                                    # 获取科目ID
                                    row_id = await row.get_attribute("id")
                                    
                                    # 检查是否有可编辑的金额输入框
                                    amount_input = await row.locator("input[cname='t.value']").first
                                    if await amount_input.is_visible() and not await amount_input.get_attribute("readonly"):
                                        # 获取科目描述
                                        desc_cell = await row.locator("td[aria-describedby*='t.b_description']").first
                                        description = await desc_cell.text_content() if await desc_cell.is_visible() else ""
                                        description = description.strip() if description else ""
                                        
                                        subjects.append({
                                            "id": row_id,
                                            "name": subject_name,
                                            "description": description,
                                            "input_selector": f"#{await amount_input.get_attribute('id')}"
                                        })
                                        logger.info(f"找到可编辑科目: {subject_name} (ID: {row_id})")
                            except Exception as e:
                                logger.debug(f"处理科目行时出错: {e}")
                                continue
            
            logger.info(f"总共找到 {len(subjects)} 个可编辑的预约科目")
            return subjects
            
        except Exception as e:
            logger.error(f"获取预约科目信息失败: {e}")
            return []

    async def analyze_with_llm(self, appointment_subject: str, amount: float, subjects_info: list, subject_mapping: dict) -> dict:
        """使用LLM分析并确定最合适的科目"""
        try:
            # 构建提示词
            subjects_text = "\n".join([
                f"- {subject['name']}: {subject['description']}" 
                for subject in subjects_info
            ])
            
            # 构建科目映射表信息（包含含义说明）
            mapping_text = "\n".join([
                f"- {subject}: {mapping_info['input_id']} (说明: {mapping_info.get('description', '无说明')})" 
                for subject, mapping_info in subject_mapping.items()
            ])
            
            prompt = f"""
请根据以下信息，选择最合适的预约科目来填写金额：

**报销信息：**
- 预约科目: {appointment_subject}
- 金额: {amount}

**可用的预约科目列表：**
{subjects_text}

**科目-输入框ID对应表：**
{mapping_text}

请分析预约科目"{appointment_subject}"与上述科目列表的匹配度，选择最合适的一个科目。
同时考虑科目映射表中的对应关系，确保选择的科目在映射表中存在对应的输入框ID。
只返回科目名称，不要其他解释。
"""
            
            # 调用Ollama API
            response = await self.call_ollama_api(prompt)
            if not response:
                logger.error("LLM API调用失败")
                return None
            
            # 解析响应，找到匹配的科目
            logger.info(f"LLM响应: {response}")
            
            # 首先尝试精确匹配
            for subject in subjects_info:
                if subject["name"] in response:
                    logger.info(f"LLM精确匹配到科目: {subject['name']}")
                    return subject
            
            # 如果精确匹配失败，尝试模糊匹配
            logger.info("尝试模糊匹配...")
            for subject in subjects_info:
                # 检查科目名称是否包含在LLM响应中
                if subject["name"] in response or any(word in response for word in subject["name"].split()):
                    logger.info(f"LLM模糊匹配到科目: {subject['name']}")
                    return subject
            
            # 如果仍然没有匹配，尝试在映射表中查找相似的科目
            logger.info("尝试在映射表中查找相似科目...")
            for mapped_subject, mapping_info in subject_mapping.items():
                if any(word in response for word in mapped_subject.split()):
                    # 创建一个新的科目对象
                    input_id = mapping_info['input_id']
                    description = mapping_info.get('description', f"从映射表匹配的科目: {mapped_subject}")
                    matched_subject = {
                        "id": f"mapped_{mapped_subject}",
                        "name": mapped_subject,
                        "description": description,
                        "input_selector": f"#{input_id}"
                    }
                    logger.info(f"LLM在映射表中匹配到科目: {mapped_subject}, 使用输入框ID: {input_id}")
                    return matched_subject
            
            # 如果还是没有匹配，尝试直接根据预约科目名称匹配
            logger.info("尝试直接根据预约科目名称匹配...")
            for mapped_subject, mapping_info in subject_mapping.items():
                if appointment_subject in mapped_subject or mapped_subject in appointment_subject:
                    input_id = mapping_info['input_id']
                    description = mapping_info.get('description', f"根据预约科目直接匹配: {mapped_subject}")
                    matched_subject = {
                        "id": f"mapped_{mapped_subject}",
                        "name": mapped_subject,
                        "description": description,
                        "input_selector": f"#{input_id}"
                    }
                    logger.info(f"根据预约科目直接匹配到科目: {mapped_subject}, 使用输入框ID: {input_id}")
                    return matched_subject
            
            logger.warning(f"LLM响应 '{response}' 未匹配到任何科目")
            return None
            
        except Exception as e:
            logger.error(f"LLM分析失败: {e}")
            return None

    async def call_ollama_api(self, prompt: str) -> str:
        """调用Ollama API"""
        try:
            url = "http://localhost:11434/api/generate"
            data = {
                "model": "llama2",
                "prompt": prompt,
                "stream": False
            }
            
            response = requests.post(url, json=data, timeout=30)
            if response.status_code == 200:
                result = response.json()
                return result.get("response", "").strip()
            else:
                logger.error(f"Ollama API调用失败: {response.status_code}")
                return ""
                
        except Exception as e:
            logger.error(f"调用Ollama API失败: {e}")
            return ""

    async def fill_amount_to_subject(self, target_subject: dict, amount: float) -> bool:
        """填写金额到指定的科目"""
        try:
            logger.info(f"填写金额 {amount} 到科目: {target_subject['name']}")
            
            # 在iframe中查找并填写金额
            frames = self.page.frames
            for frame_idx, frame in enumerate(frames):
                if frame.url and "WF_YB6" in frame.url:
                    try:
                        # 使用科目ID查找输入框
                        input_selector = target_subject["input_selector"]
                        logger.info(f"尝试使用选择器: {input_selector}")
                        
                        # 尝试多种选择器策略
                        selectors_to_try = [
                            input_selector,  # 原始选择器
                            input_selector.replace("#", ""),  # 不带#的ID
                            f"input[id='{input_selector.replace('#', '')}']",  # 完整的input选择器
                            f"input[name='{input_selector.replace('#', '')}']",  # 使用name属性
                            f"input[type='text']",  # 通用文本输入框
                            f"input[type='number']",  # 数字输入框
                        ]
                        
                        for selector in selectors_to_try:
                            try:
                                elements = await frame.locator(selector).all()
                                for element in elements:
                                    try:
                                        is_visible = await element.is_visible()
                                        is_enabled = await element.is_enabled()
                                        
                                        if is_visible and is_enabled:
                                            # 清空输入框并填写金额
                                            await element.fill("")
                                            await element.fill(str(amount))
                                            logger.info(f"✓ 成功填写金额 {amount} 到 {target_subject['name']} (使用选择器: {selector})")
                                            return True
                                        else:
                                            logger.debug(f"输入框不可见或不可用: {selector}")
                                    except Exception as e:
                                        logger.debug(f"填写元素失败: {e}")
                                        continue
                                        
                            except Exception as e:
                                logger.debug(f"选择器 {selector} 失败: {e}")
                                continue
                        
                        # 如果所有选择器都失败，尝试手动查找所有输入框
                        logger.info("尝试手动查找所有输入框...")
                        input_elements = await frame.locator("input").all()
                        for i, element in enumerate(input_elements):
                            try:
                                name = await element.get_attribute("name") or ""
                                element_id = await element.get_attribute("id") or ""
                                is_visible = await element.is_visible()
                                is_enabled = await element.is_enabled()
                                
                                logger.info(f"输入框 {i+1}: name='{name}', id='{element_id}', visible={is_visible}, enabled={is_enabled}")
                                
                                # 如果找到看起来像金额输入框的元素，尝试填写
                                if is_visible and is_enabled and (name or element_id):
                                    await element.fill("")
                                    await element.fill(str(amount))
                                    logger.info(f"✓ 手动填写金额成功: {amount}")
                                    return True
                            except Exception as e:
                                logger.debug(f"手动填写输入框 {i+1} 失败: {e}")
                                continue
                        
                        logger.warning(f"所有方法都失败，无法找到科目 {target_subject['name']} 的输入框")
                        
                    except Exception as e:
                        logger.debug(f"在iframe {frame_idx} 中填写金额失败: {e}")
                        continue
            
            logger.error("未找到可用的金额输入框")
            return False
            
        except Exception as e:
            logger.error(f"填写金额到科目失败: {e}")
            return False
    
    async def debug_input_fields(self):
        """调试：显示页面上所有输入框信息"""
        try:
            logger.info("=== 调试：分析页面输入框 ===")
            
            # 在主页面查找所有输入框
            input_elements = await self.page.locator("input").all()
            logger.info(f"主页面找到 {len(input_elements)} 个输入框")
            
            for i, element in enumerate(input_elements):
                try:
                    name = await element.get_attribute("name") or "无"
                    element_id = await element.get_attribute("id") or "无"
                    placeholder = await element.get_attribute("placeholder") or "无"
                    title = await element.get_attribute("title") or "无"
                    input_type = await element.get_attribute("type") or "text"
                    is_visible = await element.is_visible()
                    is_enabled = await element.is_enabled()
                    
                    logger.info(f"输入框 {i+1}: name='{name}', id='{element_id}', placeholder='{placeholder}', title='{title}', type='{input_type}', visible={is_visible}, enabled={is_enabled}")
                except Exception as e:
                    logger.debug(f"获取输入框 {i+1} 信息失败: {e}")
            
            # 在主页面查找所有下拉框
            select_elements = await self.page.locator("select").all()
            logger.info(f"主页面找到 {len(select_elements)} 个下拉框")
            
            for i, element in enumerate(select_elements):
                try:
                    name = await element.get_attribute("name") or "无"
                    element_id = await element.get_attribute("id") or "无"
                    title = await element.get_attribute("title") or "无"
                    is_visible = await element.is_visible()
                    is_enabled = await element.is_enabled()
                    
                    logger.info(f"下拉框 {i+1}: name='{name}', id='{element_id}', title='{title}', visible={is_visible}, enabled={is_enabled}")
                    
                    # 获取下拉框的选项
                    options = await element.locator("option").all()
                    logger.info(f"  选项数量: {len(options)}")
                    for j, option in enumerate(options):
                        try:
                            option_value = await option.get_attribute("value") or "无"
                            option_text = await option.text_content() or "无"
                            logger.info(f"    选项 {j+1}: value='{option_value}', text='{option_text}'")
                        except Exception as e:
                            logger.debug(f"获取选项 {j+1} 信息失败: {e}")
                            
                except Exception as e:
                    logger.debug(f"获取下拉框 {i+1} 信息失败: {e}")
            
            # 在iframe中查找所有输入框和下拉框
            frames = self.page.frames
            for frame_idx, frame in enumerate(frames):
                if frame.url and "WF_YB6" in frame.url:
                    logger.info(f"=== 在iframe {frame_idx} 中查找输入框和下拉框 ===")
                    input_elements = await frame.locator("input").all()
                    logger.info(f"iframe {frame_idx} 找到 {len(input_elements)} 个输入框")
                    
                    for i, element in enumerate(input_elements):
                        try:
                            name = await element.get_attribute("name") or "无"
                            element_id = await element.get_attribute("id") or "无"
                            placeholder = await element.get_attribute("placeholder") or "无"
                            title = await element.get_attribute("title") or "无"
                            input_type = await element.get_attribute("type") or "text"
                            is_visible = await element.is_visible()
                            is_enabled = await element.is_enabled()
                            
                            logger.info(f"iframe输入框 {i+1}: name='{name}', id='{element_id}', placeholder='{placeholder}', title='{title}', type='{input_type}', visible={is_visible}, enabled={is_enabled}")
                        except Exception as e:
                            logger.debug(f"获取iframe输入框 {i+1} 信息失败: {e}")
                    
                    # 在iframe中查找所有下拉框
                    select_elements = await frame.locator("select").all()
                    logger.info(f"iframe {frame_idx} 找到 {len(select_elements)} 个下拉框")
                    
                    for i, element in enumerate(select_elements):
                        try:
                            name = await element.get_attribute("name") or "无"
                            element_id = await element.get_attribute("id") or "无"
                            title = await element.get_attribute("title") or "无"
                            is_visible = await element.is_visible()
                            is_enabled = await element.is_enabled()
                            
                            logger.info(f"iframe下拉框 {i+1}: name='{name}', id='{element_id}', title='{title}', visible={is_visible}, enabled={is_enabled}")
                            
                            # 获取下拉框的选项
                            options = await element.locator("option").all()
                            logger.info(f"  选项数量: {len(options)}")
                            for j, option in enumerate(options):
                                try:
                                    option_value = await option.get_attribute("value") or "无"
                                    option_text = await option.text_content() or "无"
                                    logger.info(f"    选项 {j+1}: value='{option_value}', text='{option_text}'")
                                except Exception as e:
                                    logger.debug(f"获取选项 {j+1} 信息失败: {e}")
                                    
                        except Exception as e:
                            logger.debug(f"获取iframe下拉框 {i+1} 信息失败: {e}")
            
        except Exception as e:
            logger.error(f"调试输入框失败: {e}")
    
    async def fill_project_number(self, project_number: str) -> bool:
        """填写项目编号"""
        try:
            logger.info(f"尝试填写项目编号: {project_number}")
            
            # 使用精确的项目编号输入框选择器
            project_selectors = [
                "input[id='formWF_YB6_230_yta-uni_prj_code']",
                "input[name='formWF_YB6_230_yta-uni_prj_code']",
                "input[name*='项目编号']",
                "input[name*='project']",
                "input[id*='项目编号']",
                "input[id*='project']",
                "input[placeholder*='项目编号']",
                "input[placeholder*='请输入项目编号']",
                "input[title*='项目编号']",
                "input[title*='project']",
                # 排除附件张数相关的选择器
                "input[name*='项目']:not([name*='附件']):not([name*='张数'])",
                "input[id*='项目']:not([id*='附件']):not([id*='张数'])",
                "input[placeholder*='项目']:not([placeholder*='附件']):not([placeholder*='张数'])",
                # 添加更多通用选择器
                "input[type='text']",
                "input:not([type='hidden']):not([type='submit']):not([type='button'])"
            ]
            
            logger.info(f"将尝试 {len(project_selectors)} 个选择器来填写项目编号")
            result = await self.fill_input_field(project_selectors, project_number, "项目编号")
            
            if not result:
                logger.warning("所有选择器都失败了，尝试手动查找输入框...")
                # 尝试手动查找所有可见的输入框
                try:
                    frames = self.page.frames
                    for frame_idx, frame in enumerate(frames):
                        if frame.url and "WF_YB6" in frame.url:
                            logger.info(f"在iframe {frame_idx} 中手动查找输入框...")
                            input_elements = await frame.locator("input").all()
                            for i, element in enumerate(input_elements):
                                try:
                                    name = await element.get_attribute("name") or ""
                                    element_id = await element.get_attribute("id") or ""
                                    placeholder = await element.get_attribute("placeholder") or ""
                                    is_visible = await element.is_visible()
                                    is_enabled = await element.is_enabled()
                                    
                                    logger.info(f"输入框 {i+1}: name='{name}', id='{element_id}', placeholder='{placeholder}', visible={is_visible}, enabled={is_enabled}")
                                    
                                    # 如果找到看起来像项目编号的输入框，尝试填写
                                    if is_visible and is_enabled and (name or element_id or placeholder):
                                        await element.fill("")
                                        await element.fill(project_number)
                                        logger.info(f"✓ 手动填写项目编号成功: {project_number}")
                                        return True
                                except Exception as e:
                                    logger.debug(f"手动填写输入框 {i+1} 失败: {e}")
                                    continue
                except Exception as e:
                    logger.error(f"手动查找输入框失败: {e}")
            
            return result
            
        except Exception as e:
            logger.error(f"填写项目编号失败: {e}")
            return False
    
    async def fill_attachment_count(self, attachment_count: int) -> bool:
        """填写附件张数"""
        try:
            logger.info(f"尝试填写附件张数: {attachment_count}")
            
            # 使用精确的附件张数输入框选择器
            attachment_selectors = [
                "input[id='formWF_YB6_230_yta-addition']",
                "input[name='formWF_YB6_230_yta-addition']",
                "input[name*='附件张数']",
                "input[name*='附件数量']",
                "input[name*='张数']",
                "input[id*='附件张数']",
                "input[id*='附件数量']",
                "input[id*='张数']",
                "input[placeholder*='附件张数']",
                "input[placeholder*='请输入附件张数']",
                "input[placeholder*='张数']",
                "input[title*='附件张数']",
                "input[title*='张数']",
                "input[type='number']",
                # 更通用的选择器，但排除项目编号相关的
                "input[name*='附件']:not([name*='项目'])",
                "input[id*='附件']:not([id*='项目'])",
                "input[placeholder*='附件']:not([placeholder*='项目'])"
            ]
            
            return await self.fill_input_field(attachment_selectors, str(attachment_count), "附件张数")
            
        except Exception as e:
            logger.error(f"填写附件张数失败: {e}")
            return False
    
    async def select_payment_method(self, payment_method: str) -> bool:
        """选择支付方式"""
        try:
            logger.info(f"尝试选择支付方式: {payment_method}")
            
            # 检查支付方式是否为空或NaN
            if not payment_method or payment_method == "nan" or payment_method.lower() == "nan":
                logger.warning("支付方式为空，跳过支付方式选择")
                return True
            
            # 支付方式映射表 - 根据HTML中的选项值
            payment_text_to_value = {
                "个人转卡": "10",
                "个人转账": "10",  # 添加个人转账映射到个人转卡
                "转账汇款": "2", 
                "合同支付": "11",
                "混合支付": "14",
                "冲销其它项目借款": "9",
                "公务卡认证还款": "15"
            }
            
            # 使用精确的支付方式下拉框选择器
            payment_selectors = [
                "select[id='formWF_YB6_230_yta-pay_type']",
                "select[name='formWF_YB6_230_yta-pay_type']",
                "select[name*='支付']",
                "select[name*='方式']",
                "select[id*='支付']",
                "select[id*='方式']",
                "select[title*='支付']",
                "select[title*='方式']",
                "select"
            ]
            
            # 尝试通过value选择
            if payment_method in payment_text_to_value:
                value = payment_text_to_value[payment_method]
                logger.info(f"使用映射值选择支付方式: {payment_method} -> {value}")
                
                # 直接使用page.select_option方法
                try:
                    # 在主页面尝试
                    await self.page.select_option("select[id='formWF_YB6_230_yta-pay_type']", value=value)
                    logger.info(f"✓ 成功选择支付方式: {payment_method} (value={value})")
                    return True
                except Exception as e1:
                    logger.debug(f"主页面选择失败: {e1}")
                    
                    # 在iframe中尝试
                    frames = self.page.frames
                    for frame in frames:
                        if frame.url and "WF_YB6" in frame.url:
                            try:
                                await frame.select_option("select[id='formWF_YB6_230_yta-pay_type']", value=value)
                                logger.info(f"✓ 在iframe中成功选择支付方式: {payment_method} (value={value})")
                                return True
                            except Exception as e2:
                                logger.debug(f"iframe选择失败: {e2}")
                                continue
                
                # 如果直接选择失败，尝试通用方法
                success = await self.select_dropdown_option_by_value(payment_selectors, value, "支付方式")
                if success:
                    return True
                else:
                    logger.warning(f"通过value选择失败，尝试通过文本选择")
            
            # 如果映射失败，尝试直接选择
            logger.info(f"尝试直接选择支付方式: {payment_method}")
            return await self.select_dropdown_option(payment_selectors, payment_method, "支付方式")
            
        except Exception as e:
            logger.error(f"选择支付方式失败: {e}")
            return False
    
    async def click_next_button(self) -> bool:
        """点击下一步按钮"""
        try:
            logger.info("=== 点击下一步按钮 ===")
            
            # 等待页面加载
            await asyncio.sleep(2)
            
            # 下一步按钮的选择器列表
            next_button_selectors = [
                "button[guid='0B0E662420BC4914918B653A17663C5F']",  # 用户提供的特定按钮
                "button[btnname='下一步']",
                "button[guid='0D08843AA61A4D22AD573C7166521CA6']",
                "button.winBtn.funcButton",
                "button:has-text('下一步')",
                "button:has-text('Next')",
                "button[title*='下一步']",
                "button[onclick*='next']",
                "button[onclick*='下一步']"
            ]
            
            # 在主页面查找下一步按钮
            for selector in next_button_selectors:
                try:
                    elements = await self.page.locator(selector).all()
                    for element in elements:
                        try:
                            is_visible = await element.is_visible()
                            is_enabled = await element.is_enabled()
                            
                            if is_visible and is_enabled:
                                logger.info(f"找到下一步按钮，使用选择器: {selector}")
                                await element.click()
                                logger.info("✓ 成功点击下一步按钮")
                                return True
                        except Exception as e:
                            logger.debug(f"点击元素失败: {e}")
                            continue
                        
                except Exception as e:
                    logger.debug(f"选择器 {selector} 失败: {e}")
                    continue
            
            # 在iframe中查找下一步按钮
            frames = self.page.frames
            for frame_idx, frame in enumerate(frames):
                if frame.url and "WF_YB6" in frame.url:
                    logger.info(f"在iframe {frame_idx} 中查找下一步按钮")
                    
                    for selector in next_button_selectors:
                        try:
                            elements = await frame.locator(selector).all()
                            for element in elements:
                                try:
                                    is_visible = await element.is_visible()
                                    is_enabled = await element.is_enabled()
                                    
                                    if is_visible and is_enabled:
                                        logger.info(f"在iframe中找到下一步按钮，使用选择器: {selector}")
                                        await element.click()
                                        logger.info("✓ 在iframe中成功点击下一步按钮")
                                        return True
                                except Exception as e:
                                    logger.debug(f"iframe中点击元素失败: {e}")
                                    continue
                                
                        except Exception as e:
                            logger.debug(f"iframe中选择器 {selector} 失败: {e}")
                            continue
            
            logger.warning("未找到下一步按钮")
            return False
            
        except Exception as e:
            logger.error(f"点击下一步按钮失败: {e}")
            return False
    
    async def fill_input_field(self, selectors: list, value: str, field_name: str) -> bool:
        """通用输入框填写函数"""
        try:
            # 在主页面查找
            for selector in selectors:
                try:
                    elements = await self.page.locator(selector).all()
                    for element in elements:
                        try:
                            is_visible = await element.is_visible()
                            is_enabled = await element.is_enabled()
                            
                            if is_visible and is_enabled:
                                await element.fill("")
                                await element.fill(value)
                                logger.info(f"✓ 成功填写{field_name}: {value} (选择器: {selector})")
                                return True
                        except Exception as e:
                            logger.debug(f"填写元素失败: {e}")
                            continue
                        
                except Exception as e:
                    logger.debug(f"选择器 {selector} 失败: {e}")
                    continue
            
            # 在iframe中查找
            frames = self.page.frames
            for frame in frames:
                if frame.url and "WF_YB6" in frame.url:
                    for selector in selectors:
                        try:
                            elements = await frame.locator(selector).all()
                            for element in elements:
                                try:
                                    is_visible = await element.is_visible()
                                    is_enabled = await element.is_enabled()
                                    
                                    if is_visible and is_enabled:
                                        await element.fill("")
                                        await element.fill(value)
                                        logger.info(f"✓ 在iframe中成功填写{field_name}: {value} (选择器: {selector})")
                                        return True
                                except Exception as e:
                                    logger.debug(f"iframe中填写元素失败: {e}")
                                    continue
                            
                        except Exception as e:
                            logger.debug(f"iframe中选择器 {selector} 失败: {e}")
                            continue
            
            logger.warning(f"未找到{field_name}输入框")
            return False
            
        except Exception as e:
            logger.error(f"填写{field_name}失败: {e}")
            return False
    
    async def select_dropdown_option(self, selectors: list, option_text: str, field_name: str) -> bool:
        """通用下拉框选择函数"""
        try:
            # 在主页面查找
            for selector in selectors:
                try:
                    elements = await self.page.locator(selector).all()
                    for element in elements:
                        try:
                            is_visible = await element.is_visible()
                            is_enabled = await element.is_enabled()
                            
                            if is_visible and is_enabled:
                                # 尝试多种方式选择选项
                                try:
                                    # 首先尝试通过label选择
                                    await element.select_option(label=option_text)
                                    logger.info(f"✓ 成功选择{field_name}: {option_text} (通过label, 选择器: {selector})")
                                    return True
                                except Exception as e1:
                                    logger.debug(f"通过label选择失败: {e1}")
                                    try:
                                        # 尝试通过text选择
                                        await element.select_option(text=option_text)
                                        logger.info(f"✓ 成功选择{field_name}: {option_text} (通过text, 选择器: {selector})")
                                        return True
                                    except Exception as e2:
                                        logger.debug(f"通过text选择失败: {e2}")
                                        try:
                                            # 尝试通过value选择
                                            await element.select_option(value=option_text)
                                            logger.info(f"✓ 成功选择{field_name}: {option_text} (通过value, 选择器: {selector})")
                                            return True
                                        except Exception as e3:
                                            logger.debug(f"通过value选择失败: {e3}")
                                            # 最后尝试点击下拉框然后选择
                                            try:
                                                await element.click()
                                                await asyncio.sleep(1)
                                                # 查找包含指定文本的option
                                                option_locator = element.locator(f"option:has-text('{option_text}')")
                                                await option_locator.click()
                                                logger.info(f"✓ 成功选择{field_name}: {option_text} (通过点击, 选择器: {selector})")
                                                return True
                                            except Exception as e4:
                                                logger.debug(f"通过点击选择失败: {e4}")
                                                continue
                        except Exception as e:
                            logger.debug(f"选择下拉框选项失败: {e}")
                            continue
                        
                except Exception as e:
                    logger.debug(f"选择器 {selector} 失败: {e}")
                    continue
            
            # 在iframe中查找
            frames = self.page.frames
            for frame in frames:
                if frame.url and "WF_YB6" in frame.url:
                    for selector in selectors:
                        try:
                            elements = await frame.locator(selector).all()
                            for element in elements:
                                try:
                                    is_visible = await element.is_visible()
                                    is_enabled = await element.is_enabled()
                                    
                                    if is_visible and is_enabled:
                                        # 尝试多种方式选择选项
                                        try:
                                            # 首先尝试通过label选择
                                            await element.select_option(label=option_text)
                                            logger.info(f"✓ 在iframe中成功选择{field_name}: {option_text} (通过label, 选择器: {selector})")
                                            return True
                                        except Exception as e1:
                                            logger.debug(f"iframe中通过label选择失败: {e1}")
                                            try:
                                                # 尝试通过text选择
                                                await element.select_option(text=option_text)
                                                logger.info(f"✓ 在iframe中成功选择{field_name}: {option_text} (通过text, 选择器: {selector})")
                                                return True
                                            except Exception as e2:
                                                logger.debug(f"iframe中通过text选择失败: {e2}")
                                                try:
                                                    # 尝试通过value选择
                                                    await element.select_option(value=option_text)
                                                    logger.info(f"✓ 在iframe中成功选择{field_name}: {option_text} (通过value, 选择器: {selector})")
                                                    return True
                                                except Exception as e3:
                                                    logger.debug(f"iframe中通过value选择失败: {e3}")
                                                    # 最后尝试点击下拉框然后选择
                                                    try:
                                                        await element.click()
                                                        await asyncio.sleep(1)
                                                        # 查找包含指定文本的option
                                                        option_locator = element.locator(f"option:has-text('{option_text}')")
                                                        await option_locator.click()
                                                        logger.info(f"✓ 在iframe中成功选择{field_name}: {option_text} (通过点击, 选择器: {selector})")
                                                        return True
                                                    except Exception as e4:
                                                        logger.debug(f"iframe中通过点击选择失败: {e4}")
                                                        continue
                                except Exception as e:
                                    logger.debug(f"iframe中选择下拉框选项失败: {e}")
                                    continue
                            
                        except Exception as e:
                            logger.debug(f"iframe中选择器 {selector} 失败: {e}")
                            continue
            
            logger.warning(f"未找到{field_name}下拉框或选项")
            return False
            
        except Exception as e:
            logger.error(f"选择{field_name}失败: {e}")
            return False
    

        """通过value值选择下拉框选项"""
        try:
            # 在主页面查找
            for selector in selectors:
                try:
                    elements = await self.page.locator(selector).all()
                    for element in elements:
                        try:
                            is_visible = await element.is_visible()
                            is_enabled = await element.is_enabled()
                            
                            if is_visible and is_enabled:
                                # 通过value选择
                                await element.select_option(value=value)
                                logger.info(f"✓ 成功选择{field_name}: value={value} (选择器: {selector})")
                                return True
                        except Exception as e:
                            logger.debug(f"通过value选择失败: {e}")
                            continue
                        
                except Exception as e:
                    logger.debug(f"选择器 {selector} 失败: {e}")
                    continue
            
            # 在iframe中查找
            frames = self.page.frames
            for frame in frames:
                if frame.url and "WF_YB6" in frame.url:
                    for selector in selectors:
                        try:
                            elements = await frame.locator(selector).all()
                            for element in elements:
                                try:
                                    is_visible = await element.is_visible()
                                    is_enabled = await element.is_enabled()
                                    
                                    if is_visible and is_enabled:
                                        # 通过value选择
                                        await element.select_option(value=value)
                                        logger.info(f"✓ 在iframe中成功选择{field_name}: value={value} (选择器: {selector})")
                                        return True
                                except Exception as e:
                                    logger.debug(f"iframe中通过value选择失败: {e}")
                                    continue
                            
                        except Exception as e:
                            logger.debug(f"iframe中选择器 {selector} 失败: {e}")
                            continue
            
            logger.warning(f"未找到{field_name}下拉框或选项")
            return False
            
        except Exception as e:
            logger.error(f"通过value选择{field_name}失败: {e}")
            return False
        
    def add_expense(self, expense: ExpenseItem) -> None:
        """添加报销项目"""
        self.expenses.append(expense)
        logger.info(f"已添加报销项目: {expense.project} - {expense.account} - ¥{expense.amount}")
    
    async def start_browser(self) -> None:
        """启动浏览器"""
        self.playwright = await async_playwright().start()
        
        # 尝试使用Edge浏览器
        try:
            self.browser = await self.playwright.chromium.launch(
                headless=False,  # 显示浏览器窗口
                channel="msedge"  # 使用Edge浏览器
            )
            logger.info("Edge浏览器启动成功")
        except Exception as e:
            logger.warning(f"Edge浏览器启动失败: {e}")
            logger.info("使用默认浏览器...")
            self.browser = await self.playwright.chromium.launch(headless=False)
            logger.info("默认浏览器启动成功")
        
        self.page = await self.browser.new_page()
        
        # 设置用户代理
        await self.page.set_extra_http_headers({
            'User-Agent': config.BROWSER_CONFIG['user_agent']
        })
        
        
        logger.info("浏览器页面已创建")
    
    async def close_browser(self) -> None:
        """关闭浏览器"""
        if self.browser:
            await self.browser.close()
        if hasattr(self, 'playwright'):
            await self.playwright.stop()
        
        # 清理临时文件
        try:
            import os
            if os.path.exists("current_captcha.png"):
                os.remove("current_captcha.png")
                logger.info("✓ 验证码图片已清理")
        except Exception as e:
            logger.warning(f"清理验证码图片失败: {e}")
        
        logger.info("浏览器已关闭")
    
    def get_user_input(self) -> UserInputData:
        """获取用户输入信息"""
        print("\n" + "="*50)
        print("=== 报销信息输入 ===")
        print("="*50)
        
        # 获取姓名
        while True:
            name = input("请输入姓名: ").strip()
            if name:
                break
            print("姓名不能为空，请重新输入")
        
        # 获取项目编号
        while True:
            project_number = input("请输入项目编号: ").strip()
            if project_number:
                break
            print("项目编号不能为空，请重新输入")
        
        # 获取金额
        while True:
            try:
                amount_str = input("请输入金额 (元): ").strip()
                amount = float(amount_str)
                if amount > 0:
                    break
                else:
                    print("金额必须大于0，请重新输入")
            except ValueError:
                print("请输入有效的数字金额")
        
        user_data = UserInputData(name=name, project_number=project_number, amount=amount)
        
        print(f"\n✓ 输入信息确认:")
        print(f"  姓名: {user_data.name}")
        print(f"  项目编号: {user_data.project_number}")
        print(f"  金额: ¥{user_data.amount:.2f}")
        
        return user_data
    
    def get_login_credentials(self) -> tuple:
        """获取登录凭据"""
        print("\n=== 登录凭据设置 ===")
        
        # 从配置文件读取凭据
        config_username = self.config.get("username", "")
        config_password = self.config.get("password", "")
        
        # 检查配置文件中的凭据是否有效
        if config_username and config_password and config_username != "your_username" and config_password != "your_password":
            print(f"✓ 使用配置文件中的凭据: {config_username}")
            return config_username, config_password
        
        # 如果配置文件中的凭据无效，则手动输入
        print("配置文件中的凭据无效，请输入登录凭据:")
        username = input("工号: ").strip()
        password = input("密码: ").strip()
        
        return username, password
    
    async def navigate_to_login_page(self) -> bool:
        """导航到登录页面"""
        try:
            logger.info("正在访问登录页面...")
            await self.page.goto(self.config["login_url"])
            await self.page.wait_for_load_state("networkidle")
            
            # 检查是否在登录页面
            title = await self.page.title()
            logger.info(f"页面标题: {title}")
            
            # 等待登录表单加载
            await self.page.wait_for_selector(self.config["selectors"]["username_input"], timeout=10000)
            logger.info("✓ 登录页面加载成功")
            return True
            
        except Exception as e:
            logger.error(f"访问登录页面失败: {e}")
            return False
    
    async def handle_captcha(self) -> str:
        """处理验证码 - 手动输入"""
        try:
            logger.info("=== 验证码处理 ===")
            
            # 检查是否有验证码输入框
            captcha_input = self.config["selectors"]["captcha_input"]
            captcha_image = self.config["selectors"].get("captcha_image", "img[id='checkcodeImg']")
            
            # 检查验证码输入框是否存在
            captcha_exists = await self.page.locator(captcha_input).count() > 0
            if not captcha_exists:
                logger.info("✓ 无需验证码")
                return ""
            
            # 显示验证码图片信息
            try:
                captcha_element = await self.page.locator(captcha_image).first
                if captcha_element:
                    src = await captcha_element.get_attribute("src")
                    logger.info(f"验证码图片src: {src}")
                    
                    # 截图保存验证码图片供用户查看
                    captcha_screenshot_path = "current_captcha.png"
                    await captcha_element.screenshot(path=captcha_screenshot_path)
                    logger.info(f"✓ 验证码图片已保存: {captcha_screenshot_path}")
                    print(f"验证码图片已保存到: {captcha_screenshot_path}")
            except Exception as e:
                logger.warning(f"无法获取验证码图片信息: {e}")
            
            # 手动输入验证码
            logger.info("请查看浏览器中的验证码图片，然后手动输入:")
            print("\n请在浏览器中查看验证码，然后输入:")
            captcha_code = input("验证码: ").strip()
            
            # 填写验证码
            if captcha_code:
                await self.page.fill(captcha_input, captcha_code)
                logger.info(f"✓ 手动输入验证码: {captcha_code}")
                return captcha_code
            else:
                logger.warning("未输入验证码")
                return ""
                
        except Exception as e:
            logger.error(f"验证码处理失败: {e}")
            return ""
    
    async def perform_login(self, username: str, password: str) -> bool:
        """执行登录操作"""
        try:
            logger.info("=== 开始登录 ===")
            
            # 验证码处理方式：手动输入
            logger.info("验证码处理方式：手动输入")
            
            # 导航到登录页面
            if not await self.navigate_to_login_page():
                return False
            
            # 等待页面完全加载
            await asyncio.sleep(2)
            
            # 清空并填写用户名
            await self.page.fill(self.config["selectors"]["username_input"], "")
            await asyncio.sleep(0.5)
            await self.page.fill(self.config["selectors"]["username_input"], username)
            logger.info(f"✓ 用户名已填写: {username}")
            await asyncio.sleep(0.5)
            
            # 清空并填写密码
            await self.page.fill(self.config["selectors"]["password_input"], "")
            await asyncio.sleep(0.5)
            await self.page.fill(self.config["selectors"]["password_input"], password)
            logger.info("✓ 密码已填写")
            await asyncio.sleep(0.5)
            
            # 处理验证码
            captcha_code = await self.handle_captcha()
            
            # 点击登录按钮
            await self.page.click(self.config["selectors"]["login_button"])
            logger.info("✓ 已点击登录按钮")
            
            # 等待登录结果
            await asyncio.sleep(5)
            
            # 检查登录是否成功
            current_url = self.page.url
            title = await self.page.title()
            
            logger.info(f"当前URL: {current_url}")
            logger.info(f"页面标题: {title}")
            
            # 检查是否登录成功
            if "home.jsp" in current_url or "电子科技大学财务综合信息门户" in title:
                logger.info("✓ 登录成功！")
                self.is_logged_in = True
                return True
            else:
                # 检查是否有错误信息
                try:
                    error_elements = await self.page.locator(".error, .alert, .message, .errMsg").all()
                    for element in error_elements:
                        error_text = await element.text_content()
                        if error_text and error_text.strip():
                            logger.error(f"登录失败: {error_text}")
                            break
                except:
                    pass
                
                # 检查页面内容中是否有错误信息
                page_content = await self.page.content()
                if "用户名或密码错误" in page_content or "登录失败" in page_content:
                    logger.error("登录失败: 用户名或密码错误")
                elif "验证码错误" in page_content:
                    logger.error("登录失败: 验证码错误")
                else:
                    logger.error("登录失败，请检查用户名、密码和验证码")
                
                return False
                
        except Exception as e:
            logger.error(f"登录过程出错: {e}")
            return False
    
    async def navigate_to_uestc_financial(self) -> bool:
        """导航到电子科技大学财务系统"""
        try:
            logger.info("正在访问电子科技大学财务综合信息门户...")
            await self.page.goto(self.config["base_url"])
            await self.page.wait_for_load_state("networkidle")
            
            # 检查页面标题
            title = await self.page.title()
            logger.info(f"页面标题: {title}")
            
            if "电子科技大学财务综合信息门户" in title:
                logger.info("✓ 成功访问财务系统主页")
                return True
            else:
                logger.warning("页面标题不匹配，可能不是正确的财务系统页面")
                return False
                
        except Exception as e:
            logger.error(f"访问财务系统失败: {e}")
            return False
    
    async def demonstrate_system_navigation(self) -> bool:
        """演示系统导航功能"""
        try:
            logger.info("=== 演示系统导航功能 ===")
            
            # 等待系统导航区域加载
            await self.page.wait_for_selector(self.config["selectors"]["system_navigator"], timeout=10000)
            logger.info("✓ 系统导航区域已加载")
            
            # 显示欢迎信息
            try:
                welcome_text = await self.page.locator(self.config["selectors"]["welcome_message"]).text_content()
                logger.info(f"欢迎信息: {welcome_text}")
            except:
                logger.info("未找到欢迎信息")
            
            # 演示各个功能模块
            modules = self.project_config["system_modules"]
            for module_key, module_info in modules.items():
                logger.info(f"发现功能模块: {module_info['name']} ({module_info['id']})")
            
            # 尝试点击网上预约报账按钮
            logger.info("尝试点击网上预约报账按钮...")
            try:
                # 尝试点击"网上预约"链接
                await self.page.click("text=网上预约")
                await asyncio.sleep(3)
                logger.info("✓ 成功点击网上预约报账按钮")
                
                # 检查是否进入子系统
                current_url = self.page.url
                logger.info(f"当前URL: {current_url}")
                
                # 等待页面加载完成
                await asyncio.sleep(2)
                
                # 尝试点击申请报销单按钮
                logger.info("尝试点击申请报销单按钮...")
                try:
                    # 首先尝试在iframe中查找
                    frames = self.page.frames
                    button_found = False
                    
                    for frame in frames:
                        if frame.url and "WF_YB6" in frame.url:
                            logger.info(f"在iframe中查找按钮: {frame.url}")
                            try:
                                # 使用正确的选择器
                                selectors = [
                                    "button[btnname='申请报销单']",
                                    "button[guid='D02B3EF852B84C93B3245737DC749AE4']",
                                    "button.winBtn.funcButton",
                                    "text=申请报销单"
                                ]
                                
                                for selector in selectors:
                                    try:
                                        button = frame.locator(selector).first
                                        if await button.count() > 0:
                                            await button.click()
                                            logger.info(f"✓ 在iframe中成功点击申请报销单按钮 (选择器: {selector})")
                                            button_found = True
                                            break
                                    except Exception as e:
                                        logger.debug(f"选择器 {selector} 失败: {e}")
                                        continue
                                
                                if button_found:
                                    break
                            except Exception as e:
                                logger.warning(f"在iframe中查找按钮时出错: {e}")
                                continue
                    
                    # 如果iframe中没找到，尝试在主页面查找
                    if not button_found:
                        logger.info("在iframe中未找到按钮，尝试在主页面查找...")
                        selectors = [
                            "button[btnname='申请报销单']",
                            "button[guid='D02B3EF852B84C93B3245737DC749AE4']",
                            "button.winBtn.funcButton",
                            "text=申请报销单"
                        ]
                        
                        for selector in selectors:
                            try:
                                button = self.page.locator(selector).first
                                if await button.count() > 0:
                                    await button.click()
                                    logger.info(f"✓ 在主页面成功点击申请报销单按钮 (选择器: {selector})")
                                    button_found = True
                                    break
                            except Exception as e:
                                logger.debug(f"主页面选择器 {selector} 失败: {e}")
                                continue
                    
                    if button_found:
                        await asyncio.sleep(3)
                        # 检查是否进入申请页面
                        current_url = self.page.url
                        logger.info(f"申请页面URL: {current_url}")
                        logger.info("✓ 成功进入申请报销单页面，保持在当前界面")
                        
                        # 等待页面加载完成后，尝试点击"已阅读并同意"按钮
                        logger.info("等待页面加载，然后尝试点击'已阅读并同意'按钮...")
                        await asyncio.sleep(2)
                        
                        try:
                            # 在iframe中查找"已阅读并同意"按钮
                            agree_button_found = False
                            for frame in frames:
                                if frame.url and "WF_YB6" in frame.url:
                                    try:
                                        agree_selectors = [
                                            "button[btnname='已阅读并同意']",
                                            "button.winBtn.funcButton",
                                            "text=已阅读并同意"
                                        ]
                                        
                                        for selector in agree_selectors:
                                            try:
                                                agree_button = frame.locator(selector).first
                                                if await agree_button.count() > 0:
                                                    await agree_button.click()
                                                    logger.info(f"✓ 成功点击'已阅读并同意'按钮 (选择器: {selector})")
                                                    agree_button_found = True
                                                    break
                                            except Exception as e:
                                                logger.debug(f"同意按钮选择器 {selector} 失败: {e}")
                                                continue
                                        
                                        if agree_button_found:
                                            break
                                    except Exception as e:
                                        logger.warning(f"在iframe中查找同意按钮时出错: {e}")
                                        continue
                            
                            if not agree_button_found:
                                logger.warning("未找到'已阅读并同意'按钮")
                                logger.info("可能页面结构不同或按钮名称不同")
                            else:
                                # 成功点击"已阅读并同意"按钮后，填写报销表单
                                logger.info("✓ 成功点击'已阅读并同意'按钮，开始填写报销表单...")
                                await self.fill_expense_form()
                            
                        except Exception as e:
                            logger.warning(f"点击'已阅读并同意'按钮失败: {e}")
                    else:
                        logger.warning("未找到申请报销单按钮")
                        logger.info("可能页面结构不同或按钮名称不同")
                    
                except Exception as e:
                     logger.warning(f"点击申请报销单按钮失败: {e}")
                     logger.info("可能按钮选择器需要调整或页面结构不同")
                
            except Exception as e:
                logger.warning(f"点击网上预约报账按钮失败: {e}")
            
            return True
            
        except Exception as e:
            logger.error(f"系统导航演示失败: {e}")
            return False
    
    async def demonstrate_password_change(self) -> bool:
        """演示密码修改功能"""
        try:
            logger.info("=== 演示密码修改功能 ===")
            
            # 点击修改密码按钮
            await self.page.click(self.config["selectors"]["change_password_button"])
            await asyncio.sleep(1)
            
            # 等待密码修改弹窗出现
            await self.page.wait_for_selector(self.config["selectors"]["password_dialog"], timeout=5000)
            logger.info("✓ 密码修改弹窗已打开")
            
            # 填写新密码
            await self.page.fill(self.config["selectors"]["new_password1"], "test123456")
            await asyncio.sleep(0.5)
            
            await self.page.fill(self.config["selectors"]["new_password2"], "test123456")
            await asyncio.sleep(0.5)
            
            logger.info("✓ 已填写新密码")
            
            # 取消修改（不实际提交）
            await self.page.click(self.config["selectors"]["cancel_password_change"])
            await asyncio.sleep(1)
            
            logger.info("✓ 已取消密码修改")
            return True
            
        except Exception as e:
            logger.error(f"密码修改演示失败: {e}")
            return False
    
    async def demonstrate_expense_automation(self, expense: ExpenseItem) -> bool:
        """演示报销自动化流程"""
        try:
            logger.info(f"=== 演示报销自动化流程: {expense.project} ===")
            
            # 步骤1: 进入网上预约报账模块
            logger.info("步骤1: 进入网上预约报账模块")
            await self.page.click(self.config["selectors"]["online_appointment"])
            await asyncio.sleep(3)
            
            # 步骤2: 等待子系统加载
            logger.info("步骤2: 等待子系统加载")
            try:
                # 等待iframe加载
                await self.page.wait_for_selector(self.config["selectors"]["sub_system_frame"], timeout=10000)
                logger.info("✓ 子系统框架已加载")
                
                # 切换到iframe
                frame = self.page.frame_locator(self.config["selectors"]["sub_system_frame"]).first
                logger.info("✓ 已切换到子系统框架")
                
                # 这里需要根据实际的智能报销页面结构来填写表单
                # 由于没有实际的表单页面，我们模拟填写过程
                logger.info("步骤3: 模拟填写报销表单")
                logger.info(f"  - 项目: {expense.project}")
                logger.info(f"  - 科目: {expense.account}")
                logger.info(f"  - 金额: ¥{expense.amount}")
                logger.info(f"  - 描述: {expense.description}")
                logger.info(f"  - 日期: {expense.date}")
                logger.info(f"  - 类别: {expense.category}")
                
                # 模拟等待表单填写完成
                await asyncio.sleep(2)
                
                logger.info("步骤4: 模拟提交表单")
                await asyncio.sleep(1)
                
                logger.info("✓ 报销自动化流程演示完成")
                
            except Exception as e:
                logger.warning(f"子系统操作失败: {e}")
                logger.info("继续演示其他功能...")
            
            # 返回主页
            await self.page.goto(self.config["base_url"])
            await asyncio.sleep(2)
            
            return True
            
        except Exception as e:
            logger.error(f"报销自动化演示失败: {e}")
            return False
    
    def generate_summary_report(self) -> str:
        """生成汇总报告"""
        if not self.expenses:
            return "暂无报销数据"
        
        total_amount = sum(exp.amount for exp in self.expenses)
        project_summary = {}
        account_summary = {}
        
        for expense in self.expenses:
            # 按项目汇总
            if expense.project not in project_summary:
                project_summary[expense.project] = 0
            project_summary[expense.project] += expense.amount
            
            # 按科目汇总
            if expense.account not in account_summary:
                account_summary[expense.account] = 0
            account_summary[expense.account] += expense.amount
        
        report = f"""
=== 电子科技大学财务报销汇总报告 ===
系统名称: {self.project_config['system_name']}
生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
登录状态: {'已登录' if self.is_logged_in else '未登录'}
总报销金额: ¥{total_amount:.2f}
总报销项目数: {len(self.expenses)}

按项目汇总:
"""
        for project, amount in project_summary.items():
            report += f"  {project}: ¥{amount:.2f}\n"
        
        report += "\n按科目汇总:\n"
        for account, amount in account_summary.items():
            report += f"  {account}: ¥{amount:.2f}\n"
        
        return report

async def main():
    """主函数 - 电子科技大学财务系统演示"""
    print("=== 电子科技大学财务系统自动化演示 ===\n")
    
    automation = UESTCFinancialAutomation()
    
    # 添加示例报销项目（使用电子科技大学相关的项目）
    expenses_data = [
        ExpenseItem("科研项目", "差旅费", 1200.50, "参加学术会议差旅费", "2024-07-31", "住宿费", "如家酒店", "INV001"),
        ExpenseItem("教学项目", "交通费", 350.00, "教学出差高铁票", "2024-07-30", "交通费", "12306", "INV002"),
        ExpenseItem("行政项目", "办公用品", 89.90, "办公用打印纸", "2024-07-29", "办公用品", "京东", "INV003"),
        ExpenseItem("基建项目", "材料费", 156.00, "实验室材料费", "2024-07-28", "材料费", "供应商", "INV004"),
    ]
    
    print("正在添加报销项目...")
    for expense in expenses_data:
        automation.add_expense(expense)
    
    # 生成汇总报告
    print("\n正在生成汇总报告...")
    report = automation.generate_summary_report()
    print(report)
    
    # 跳过用户输入，直接进入浏览器自动化演示
    print("\n=== 浏览器自动化演示 ===")
    print("注意：即将打开浏览器窗口进行登录演示...")
    print("跳过用户输入，直接进入自动化演示...")
    
    try:
        # 启动浏览器
        await automation.start_browser()
        
        # 获取登录凭据
        username, password = automation.get_login_credentials()
        
        if not username or not password:
            print("未提供登录凭据，将进行只读演示...")
            # 进行只读演示
            if await automation.navigate_to_uestc_financial():
                # 演示系统导航
                await automation.demonstrate_system_navigation()
                print("\n✓ 只读演示完成！")
            else:
                print("\n✗ 无法访问财务系统")
        else:
            # 执行登录
            if await automation.perform_login(username, password):
                # 演示登录后功能
                await automation.demonstrate_system_navigation()
                print("\n✓ 所有演示功能完成！")
            else:
                print("\n✗ 登录失败，无法演示登录后功能")
        
        # 等待用户确认后关闭浏览器
        input("\n按回车键关闭浏览器...")
        
    except Exception as e:
        print(f"\n✗ 浏览器自动化失败: {e}")
    
    finally:
        # 关闭浏览器
        await automation.close_browser()
    
    print("\n=== 演示完成 ===")
    print("这个演示展示了如何自动化登录和操作电子科技大学财务综合信息门户")
    print("在实际使用中，请注意:")
    print("1. 确保有权限使用自动化工具访问系统")
    print("2. 遵守学校的使用政策和规定")
    print("3. 妥善保管登录凭据")
    print("4. 定期更新密码")

if __name__ == "__main__":
    # 运行主函数
    asyncio.run(main()) 