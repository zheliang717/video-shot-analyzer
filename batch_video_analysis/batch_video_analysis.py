import os
import cv2
import pandas as pd
import dashscope
import tempfile
from PIL import Image
import re
import json
import numpy as np

# ========== 配置区 ==========
VIDEO_DIR = r'D:/Desktop/视频分析工具/PR分割导出'   # ← 修改为你的视频目录
OUTPUT_CSV = 'analysis_output.csv'                 # 输出文件名
STILL_FRAME_DIR = 'stills'                         # 定帧图片输出目录
DASHSCOPE_API_KEY = ''                  # ← 填你的API-KEY
MODEL_NAME = 'qwen-vl-plus'                       # 模型名称，可自行更换
# ========== END =============

os.makedirs(STILL_FRAME_DIR, exist_ok=True)
dashscope.api_key = DASHSCOPE_API_KEY

STRUCT_PROMPT = """
你是一名专业影视分镜分析师。请只输出如下JSON结构分镜信息（不要解释或说明，只能有如下内容）：

{
  "景别": "",
  "焦段": "",
  "运镜": "",
  "机位": "",
  "画面": "",
  "场景": ""
}

请严格遵循以下规范：
1. “运镜”只能从如下类型及方向中选择，格式示例：左移、右摇、左跟、右环、上升、下降、左转、右旋转、固定等。若无方向则写“不确定方向”。可选项：推、拉、移、摇、跟、环、转（旋转）、升降、固定。
2. “机位”只能为“低机位”“中机位”“高机位”三选一，不可填其他，如无法判断只能写“不确定”。
3. “焦段”只能填“广角”“短焦”“中焦”“长焦”四选一。如无法判断可按景别推测或写“不确定”。
4. 其他项按画面简明描述。
请严格按上述要求输出JSON，每项都必须填写，且只能用指定内容，否则填“不确定”。
"""

# ---------- 标准化工具 ----------
def standardize_focal(focal, desc):
    # 归一化为广角/短焦/中焦/长焦
    if '广角' in focal:
        return '广角'
    if '短焦' in focal:
        return '短焦'
    if '中焦' in focal:
        return '中焦'
    if '长焦' in focal:
        return '长焦'
    # 识别常见焦距数字
    m = re.search(r'(\d{2,3})mm', focal)
    if m:
        num = int(m.group(1))
        if num <= 24:
            return '广角'
        elif 25 <= num <= 35:
            return '短焦'
        elif 36 <= num <= 70:
            return '中焦'
        elif num >= 85:
            return '长焦'
    # 根据景别推断
    scene = desc.get('景别', '')
    if any(x in scene for x in ['全景', '远景', '广阔', '大空间']):
        return '广角'
    elif any(x in scene for x in ['近景', '特写']):
        return '中焦'
    elif '中景' in scene:
        return '短焦'
    return '不确定'

def standardize_move(move):
    # 可选类型和方向
    main_types = ['移', '推', '拉', '摇', '跟', '环', '转', '旋转', '升', '降', '固定']
    directions = ['左', '右', '上', '下', '前', '后', '顺时针', '逆时针']
    move = move.replace(' ', '').replace('(', '').replace(')', '')
    for t in main_types:
        if t in move:
            # 匹配方向
            for d in directions:
                if d in move:
                    return d + t
            # 固定、升降等特殊
            if t == '固定':
                return '固定'
            if t in ['升', '降']:
                return t
            return t + '(不确定方向)'
    return '不确定'

def standardize_angle(angle):
    allowed = ['低机位','中机位','高机位']
    return angle if angle in allowed else '不确定'

def fix_motion_if_possible(desc, motion_flag):
    # AI若“固定”，且实际检测有运动，则兜底补“移(不确定方向)”
    if desc.get('运镜', '') == '固定' and motion_flag:
        return '移(不确定方向)'
    return desc.get('运镜', '')

# ----------- 帧差判别运动 -----------
def detect_camera_motion(frame1, frame2, threshold=8.0):
    if frame1 is None or frame2 is None:
        return False
    h = min(frame1.shape[0], frame2.shape[0])
    w = min(frame1.shape[1], frame2.shape[1])
    f1 = cv2.resize(frame1, (w, h))
    f2 = cv2.resize(frame2, (w, h))
    gray1 = cv2.cvtColor(f1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(f2, cv2.COLOR_BGR2GRAY)
    diff = cv2.absdiff(gray1, gray2)
    mean_diff = np.mean(diff)
    return mean_diff > threshold

# ------------- AI结构化分镜 -------------
def get_structured_description(frame, model=MODEL_NAME):
    img_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(img_rgb)
    tmp = tempfile.NamedTemporaryFile(suffix='.jpg', delete=False)
    try:
        pil_img.save(tmp, format='JPEG')
        tmp.close()
        image_path = tmp.name

        result = dashscope.MultiModalConversation.call(
            model=model,
            messages=[
                {"role": "system", "content": "你是专业影视剧分镜分析师。"},
                {"role": "user", "content": [
                    {"image": image_path},
                    {"text": STRUCT_PROMPT}
                ]}
            ]
        )
        content = result['output']['choices'][0]['message']['content']

        if isinstance(content, dict):
            if 'text' in content:
                content = content['text']
            else:
                content = str(content)
        if isinstance(content, list):
            new_content = ''
            for x in content:
                if isinstance(x, dict) and 'text' in x:
                    new_content += x['text']
                else:
                    new_content += str(x)
            content = new_content

        m = re.search(r'```json(.*?)```', content, re.S)
        if m:
            content = m.group(1)
        match = re.search(r'\{[\s\S]*?\}', content)
        if match:
            json_block = match.group()
            try:
                desc = json.loads(json_block)
            except Exception:
                json_block = json_block.replace('“', '"').replace('”', '"').replace("'", '"').replace("：", ":")
                try:
                    desc = json.loads(json_block)
                except Exception:
                    desc = {k: "" for k in ["景别", "焦段", "运镜", "机位", "画面", "场景"]}
                    for key in desc.keys():
                        m = re.search(rf'{key}\s*[:：]\s*[“"]?([^“”"\'\n,}}]*)', json_block)
                        if m:
                            desc[key] = m.group(1).strip()
        else:
            desc = {k: "" for k in ["景别", "焦段", "运镜", "机位", "画面", "场景"]}
            for key in desc.keys():
                m = re.search(rf'{key}\s*[:：]\s*[“"]?([^“”"\'\n,}}]*)', content)
                if m:
                    desc[key] = m.group(1).strip()

        desc['焦段'] = standardize_focal(desc.get('焦段', ''), desc)
        desc['运镜'] = standardize_move(desc.get('运镜',''))
        desc['机位'] = standardize_angle(desc.get('机位',''))
        return desc
    except Exception as e:
        return {
            "景别": f"生成失败：{e}",
            "焦段": "",
            "运镜": "",
            "机位": "",
            "画面": "",
            "场景": ""
        }
    finally:
        if os.path.exists(tmp.name):
            try:
                os.remove(tmp.name)
            except Exception:
                pass

def get_video_duration(video_path):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return 0
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    fps = cap.get(cv2.CAP_PROP_FPS)
    duration = frame_count / fps if fps > 0 else 0
    cap.release()
    return round(duration, 2)

def get_middle_frame(video_path):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return None
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if total_frames < 1:
        cap.release()
        return None
    mid_frame_idx = total_frames // 2
    cap.set(cv2.CAP_PROP_POS_FRAMES, mid_frame_idx)
    ret, frame = cap.read()
    cap.release()
    if ret:
        return frame
    return None

def get_first_frame(video_path):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return None
    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
    ret, frame = cap.read()
    cap.release()
    if ret:
        return frame
    return None

def analyze_videos(video_dir, output_csv):
    video_files = [f for f in os.listdir(video_dir) if f.lower().endswith(('.mp4', '.mov', '.avi', '.mkv', '.flv', '.wmv'))]
    video_files.sort()
    results = []

    for idx, video_file in enumerate(video_files, 1):
        video_path = os.path.join(video_dir, video_file)
        print(f"分析第{idx}镜: {video_file}")

        duration = get_video_duration(video_path)
        mid_frame = get_middle_frame(video_path)
        first_frame = get_first_frame(video_path)
        # 保存定帧图
        still_name = f'still_{idx}_{os.path.splitext(video_file)[0]}.jpg'
        still_path = os.path.join(STILL_FRAME_DIR, still_name)
        if first_frame is not None:
            Image.fromarray(cv2.cvtColor(first_frame, cv2.COLOR_BGR2RGB)).save(still_path)

        motion_flag = False
        if mid_frame is not None and first_frame is not None:
            motion_flag = detect_camera_motion(first_frame, mid_frame)

        if mid_frame is None:
            print(f"获取中帧失败: {video_file}")
            desc = {
                "景别": "获取中帧失败",
                "焦段": "",
                "运镜": "",
                "机位": "",
                "画面": "",
                "场景": ""
            }
        else:
            desc = get_structured_description(mid_frame)
            desc['运镜'] = fix_motion_if_possible(desc, motion_flag)
            print(f"结构化分镜信息：{desc}")

        results.append({
            '镜号': idx,
            '时长': duration,
            '景别': desc.get('景别', ''),
            '焦段': desc.get('焦段', ''),
            '运镜': desc.get('运镜', ''),
            '机位': desc.get('机位', ''),
            '画面': desc.get('画面', ''),
            '场景': desc.get('场景', ''),
            '文件名': video_file,
            '定帧图': still_path
        })

    df = pd.DataFrame(results, columns=['镜号', '时长', '景别', '焦段', '运镜', '机位', '画面', '场景', '文件名', '定帧图'])
    df.to_csv(output_csv, index=False, encoding='utf-8-sig')
    print(f"\n分析完成，已保存到 {output_csv}")

if __name__ == '__main__':
    analyze_videos(VIDEO_DIR, OUTPUT_CSV)