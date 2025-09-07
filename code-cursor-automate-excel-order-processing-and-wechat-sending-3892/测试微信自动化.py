#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试微信自动化功能
"""

import sys
import time
import platform
from py_wechat_sender.main import WeChatSender

def test_wechat_automation():
    """测试微信自动化"""
    
    print("=" * 50)
    print("微信自动化功能测试")
    print("=" * 50)
    
    # 检查系统
    if platform.system().lower() != "windows":
        print("❌ 错误：此功能仅支持Windows系统")
        return False
    
    # 检查微信是否运行
    try:
        import uiautomation as auto
        main = auto.WindowControl(searchDepth=1, ClassName="WeChatMainWndForPC")
        if not main.Exists(0.5):
            print("❌ 错误：未找到微信窗口，请先启动并登录微信PC版")
            return False
        print("✅ 微信窗口检测成功")
    except Exception as e:
        print(f"❌ 微信窗口检测失败: {e}")
        return False
    
    # 创建发送器实例
    sender = WeChatSender()
    
    # 测试消息
    test_messages = [
        ("末", "这是一条测试消息，用于验证微信自动化功能是否正常工作。\n\n如果你看到这条消息，说明自动化功能运行正常！"),
    ]
    
    print("\n开始测试发送...")
    print("⚠️ 注意：将向'末'发送测试消息")
    
    # 询问用户确认
    response = input("\n是否继续测试？(y/N): ").strip().lower()
    if response != 'y':
        print("测试已取消")
        return False
    
    # 设置进度回调
    def on_progress(msg):
        print(f"📝 {msg}")
    
    def on_finished():
        print("✅ 测试发送完成")
    
    def on_failed(err):
        print(f"❌ 测试发送失败: {err}")
    
    sender.progressed.connect(on_progress)
    sender.finished.connect(on_finished)
    sender.failed.connect(on_failed)
    
    try:
        # 开始发送测试
        sender.send(test_messages, 1.0, 1.5)
        
        # 等待完成
        time.sleep(2)
        
        print("\n" + "=" * 50)
        print("测试完成！")
        print("请检查微信中是否收到测试消息")
        print("=" * 50)
        
        return True
        
    except Exception as e:
        print(f"❌ 测试过程中出错: {e}")
        return False

if __name__ == "__main__":
    try:
        success = test_wechat_automation()
        if success:
            print("\n🎉 微信自动化测试成功！")
        else:
            print("\n💥 微信自动化测试失败！")
    except KeyboardInterrupt:
        print("\n⏹️ 测试被用户中断")
    except Exception as e:
        print(f"\n💥 测试异常: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n按回车键退出...")

