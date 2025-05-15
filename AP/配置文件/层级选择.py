import sys
from importlib import import_module

from AP.配置文件.基础参数写入 import process_data
from AP.配置文件.路径配置 import experiment_record_path, experiment_report_path


def get_valid_input(prompt, options):
    """
    提示用户输入，并验证是否为有效选项。
    """
    while True:
        choice = input(f"{prompt}（{', '.join(options)}）：")
        if choice in options:
            return choice
        else:
            print("无效选项，请重新输入。")

def handle_level(level_data):
    """
    通用的层级处理函数，根据配置动态处理用户输入。
    """
    # 提示当前层级信息
    print(level_data["title"])
    for key, value in level_data["options"].items():
        print(f"{key}. {value['label']}")

    # 获取用户选择
    choice = get_valid_input("请输入选项编号", level_data["options"].keys())
    selected_option = level_data["options"][choice]

    if "action" in selected_option:
        # 动态调用模块中的处理函数
        module_path, function_name = selected_option["action"]
        try:
            module = import_module(module_path)
            if hasattr(module, function_name):
                print(f"正在执行 {function_name} 的处理逻辑...")
                getattr(module, function_name)()  # 调用函数
            else:
                print(f"模块 {module_path} 中未找到函数 {function_name}")
        except Exception as e:
            print(f"加载或执行函数时出错：{e}")
    elif "next_level" in selected_option:
        # 进入下一层级
        handle_level(selected_option["next_level"])

# 配置层级结构
menu_structure = {
    "title": "请选择类别：",
    "options": {
        "1": {
            "label": "多媒体",
            "next_level": {
                "title": "请选择多媒体子类别：",
                "options": {
                    "1": {
                        "label": "KL项目",
                        "next_level": {
                            "title": "请选择KL项目的子类别：",
                            "options": {
                                "1": {
                                    "label": "BnO重低音",
                                    "action": ("AP.配置文件.TCL.KL.BnO重低音", "BnO_subwoofer")
                                },
                                "2": {
                                    "label": "BnO主音箱低音",
                                    "action": ("AP.配置文件.TCL.KL.BnO主音箱低音", "BnO_main_speaker_bass")
                                },
                                "3": {
                                    "label": "BnO主音箱高音",
                                    "action": ("AP.配置文件.TCL.KL.BnO主音箱高音", "BnO_main_speaker_treble")
                                },
                                "4": {
                                    "label": "G0202-000313",
                                    "action": ("AP.配置文件.TCL.KL.G0202-000313", "BnO_main_speaker_bass")
                                      },
                                "5": {"label": "非BnO主音箱低音"},
                                "6": {"label": "非BnO主音箱高音"}
                            }
                        }
                    },
                    "2": {
                           "label": "12302-500224",
                            "action": ("AP.配置文件.TCL.12302-500224", "BnO_main_speaker_bass")
                        },
                    "3": {
                           "label": "G0202-000201",
                            "action": ("AP.配置文件.TCL.G0202-000201", "BnO_main_speaker_bass")
                        },
                    "4": {
                        "label": "491",
                        "action": ("AP.配置文件.创维.491", "BnO_main_speaker_bass")
                    },
                    "5": {
                        "label": "G0202-000330(TJS6)",
                        "action": ("AP.配置文件.TCL.G0202-000330(TJS6)", "BnO_main_speaker_bass")
                    },
                    "6": {
                        "label": "G0202-000331(TJS8)",
                        "action": ("AP.配置文件.TCL.G0202-000331(TJS8)", "BnO_main_speaker_bass")
                    },
                    "7": {
                        "label": "CF10",
                        "action": ("AP.配置文件.创维.CF10", "BnO_main_speaker_bass")
                    },
                    "8": {
                        "label": "283&284",
                        "action": ("AP.配置文件.TCL.283&284", "BnO_main_speaker_bass")
                    },
                    "9": {
                        "label": "12302-500240",
                        "action": ("AP.配置文件.TCL.12302-500240", "BnO_main_speaker_bass")
                    }

                }
            }
        },
        "2": {
            "label": "车载",
            "next_level": {
                "title": "请选择车载子类别：",
                "options": {
                    "1": {
                        "label": "州伊",
                        "action": ("AP.配置文件.车载.州伊", "BnO_main_speaker_bass")
                    },
                    "2": {
                        "label": "奇瑞4寸",
                        "action": ("AP.配置文件.车载.奇瑞4寸", "BnO_main_speaker_bass")
                    },
                }
            }
        },
        "3": {
            "label": "小米",
            "next_level": {
                "title": "请选择小米子类别：",
                "options": {
                    "1": {
                        "label": "S002",
                        "action": ("AP.配置文件.小米.S002", "BnO_main_speaker_bass")
                    },
                    "2": {
                        "label": "S003",
                        "action": ("AP.配置文件.小米.S003", "BnO_main_speaker_bass")
                    },
                    "3": {
                        "label": "O32",
                        "action": ("AP.配置文件.小米.O32", "BnO_main_speaker_bass")
                    },
                }
            }
        },
        "4": {
            "label": "彩讯",
            "next_level": {
                "title": "请选择彩讯子类别：",
                "options": {
                    "1": {
                        "label": "046",
                        "action": ("AP.配置文件.彩讯.046", "BnO_main_speaker_bass")
                    },
                    "2": {
                        "label": "310100030",
                        "action": ("AP.配置文件.彩讯.310100030", "BnO_main_speaker_bass")
                    }
                }
            }
        },
        "5": {
            "label": "海信",
            "next_level": {
                "title": "请选择海信子类别：",
                "options": {
                    "1": {
                        "label": "375&376",
                        "action": ("AP.配置文件.海信.375&376", "BnO_main_speaker_bass")
                    }
                }
            }
        },
        "6": {
            "label": "惠科",
            "next_level": {
                "title": "请选择惠科子类别：",
                "options": {
                    "1": {
                        "label": "K65",
                        "action": ("AP.配置文件.惠科.K65", "BnO_main_speaker_bass")
                    }
                }
            }
        },
         "7": {
            "label": "当贝",
            "next_level": {
                "title": "请选择惠科子类别：",
                "options": {
                    "1": {
                        "label": "C3G",
                        "action": ("AP.配置文件.当贝.C3G", "BnO_main_speaker_bass")
                    }
                }
            }
        }
    }
}

if __name__ == "__main__":
    # 用户输入的ID
    user_id = input("请输入ID: ")
    # # 调用共用.py中的处理函数
    process_data(user_id, experiment_record_path, experiment_report_path)
    handle_level(menu_structure)
