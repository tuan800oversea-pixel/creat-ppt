from io import BytesIO
import streamlit as st
from PIL import Image
import uuid
import hashlib
import tempfile
import os
import time
import imagehash

# ================= 页面配置 =================
st.set_page_config(layout="wide", page_title="多图片自动生成 PPT")
st.title("多图片自动生成 PPT")

# ================= 初始化 Session State =================
def init_session():
    st.session_state.images = []
    st.session_state.processed_ids = set()
    st.session_state.page = 1
    st.session_state.ppt_bytes = None
    st.session_state.temp_duplicates = []
    # 通过 key 的变动来强制清空 file_uploader
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = str(uuid.uuid4())

if "images" not in st.session_state:
    init_session()

TMP_DIR = tempfile.gettempdir()

# ================= 核心功能：清空功能 =================
def clear_all_data():
    # 物理删除临时缩略图
    for img in st.session_state.images:
        try:
            if os.path.exists(img["thumb_path"]):
                os.remove(img["thumb_path"])
        except:
            pass
    
    # 重置状态
    st.session_state.images = []
    st.session_state.processed_ids = set()
    st.session_state.page = 1
    st.session_state.ppt_bytes = None
    st.session_state.temp_duplicates = []
    # 核心：改变 key，彻底清空上传组件的文件列表
    st.session_state.uploader_key = str(uuid.uuid4())
    st.rerun()

# ================= 核心功能：重复检测弹窗 =================
@st.dialog("发现疑似重复、缩放或变色图片")
def show_duplicate_dialog():
    st.info("系统检测到以下图片内容高度相似。您可以自由决定删除哪一张（或都保留）：")
    
    # 记录用户想要删除的 UID
    uids_to_remove = set()
    
    for idx, dup in enumerate(st.session_state.temp_duplicates):
        orig = dup['original']
        curr = dup['current']
        
        col1, col2, col3 = st.columns([4, 1, 4])
        
        with col1:
            st.image(orig['thumb_path'], caption=f"已有图片: {orig['name']}", width=180)
            if st.checkbox(f"删除这张已有项", key=f"del_orig_{idx}_{orig['uid']}"):
                uids_to_remove.add(orig['uid'])
        
        with col2:
            st.markdown("<br><br><h3 style='text-align: center;'>VS</h3>", unsafe_allow_html=True)
        
        with col3:
            st.image(curr['thumb_path'], caption=f"新上传项: {curr['name']}", width=180)
            if st.checkbox(f"删除这张新上传", key=f"del_curr_{idx}_{curr['uid']}"):
                uids_to_remove.add(curr['uid'])
        
        st.divider()

    if st.button("确认处理并关闭弹窗", type="primary", use_container_width=True):
        if uids_to_remove:
            st.session_state.images = [
                img for img in st.session_state.images 
                if img['uid'] not in uids_to_remove
            ]
        # 处理完后清空临时队列
        st.session_state.temp_duplicates = []
        st.success(f"已处理！成功删除 {len(uids_to_remove)} 张图片")
        time.sleep(0.5)
        st.rerun()

# ================= 顶部操作栏 =================
col_upload, col_clear = st.columns([8, 2])

with col_upload:
    # 使用动态 key 保证可以被 clear_all_data 彻底重置
    uploaded_files = st.file_uploader(
        "上传图片（支持批量，Ctrl+A 全选）",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=st.session_state.uploader_key
    )

with col_clear:
    st.write("---") 
    if st.button("🗑️ 清空所有图片", use_container_width=True, type="secondary"):
        clear_all_data()

# ================= 上传处理逻辑 =================
if uploaded_files:
    # 只有当临时队列为空时才进行新一轮扫描，避免重复触发
    if not st.session_state.temp_duplicates:
        new_found_duplicates = []
        # 阈值建议 15，兼顾颜色和缩放
        SIMILARITY_THRESHOLD = 15 

        for file in uploaded_files:
            file_bytes = file.read()
            file_hash = hashlib.md5(file_bytes).hexdigest()
            file_id = f"{file.name}_{file_hash}"

            if file_id in st.session_state.processed_ids:
                continue

            try:
                img = Image.open(BytesIO(file_bytes)).convert("RGB")
                curr_phash = imagehash.phash(img)
                
                uid = str(uuid.uuid4())
                thumb = img.copy()
                thumb.thumbnail((260, 260))
                thumb_path = os.path.join(TMP_DIR, f"{uid}.png")
                thumb.save(thumb_path, "PNG")

                new_img_obj = {
                    "uid": uid,
                    "name": file.name,
                    "bytes": file_bytes,
                    "thumb_path": thumb_path,
                    "ratio": img.size[0] / img.size[1],
                    "phash": curr_phash
                }

                # 查找重复
                for existing in st.session_state.images:
                    if (curr_phash - existing['phash']) <= SIMILARITY_THRESHOLD:
                        new_found_duplicates.append({"original": existing, "current": new_img_obj})
                        break # 一张图只找一个对应重复项展示
                
                st.session_state.images.append(new_img_obj)
                st.session_state.processed_ids.add(file_id)

            except Exception as e:
                st.error(f"{file.name} 读取失败：{e}")

        if new_found_duplicates:
            st.session_state.temp_duplicates = new_found_duplicates
            show_duplicate_dialog()

# ================= 展示与分页 =================
IMAGES_PER_PAGE = 40
IMAGES_PER_ROW = 10
THUMB_HEIGHT_MM = 40
MM_TO_PIXELS = 3.77953

total_images = len(st.session_state.images)
total_pages = max(1, (total_images + IMAGES_PER_PAGE - 1) // IMAGES_PER_PAGE)
st.session_state.page = min(st.session_state.page, total_pages)

start_idx = (st.session_state.page - 1) * IMAGES_PER_PAGE
page_images = st.session_state.images[start_idx : start_idx + IMAGES_PER_PAGE]

st.subheader(f"图片预览 (共 {total_images} 张)")

# 网格展示
for i in range(0, len(page_images), IMAGES_PER_ROW):
    cols = st.columns(IMAGES_PER_ROW)
    for col, img in zip(cols, page_images[i:i + IMAGES_PER_ROW]):
        with col:
            h_px = int(THUMB_HEIGHT_MM * MM_TO_PIXELS)
            w_px = int(h_px * img["ratio"])
            st.image(img["thumb_path"], width=w_px)

# 分页导航
if total_pages > 1:
    cp, cn, _ = st.columns([1, 1, 6])
    with cp:
        if st.button("上一页", disabled=(st.session_state.page <= 1)):
            st.session_state.page -= 1
            st.rerun()
    with cn:
        if st.button("下一页", disabled=(st.session_state.page >= total_pages)):
            st.session_state.page += 1
            st.rerun()

# ================= PPT 生成 =================
def generate_ppt(images):
    from pptx import Presentation
    from pptx.util import Inches, Mm
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    left_m, top_m, space, fix_h = Mm(0), Mm(10), Mm(2.5), Mm(40)
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    x, y = left_m, top_m
    
    for img in images:
        w = fix_h * img["ratio"]
        if x + w > prs.slide_width:
            x, y = left_m, y + fix_h + space
        if y + fix_h > prs.slide_height - top_m:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            x, y = left_m, top_m
        
        slide.shapes.add_picture(BytesIO(img["bytes"]), x, y, height=fix_h)
        x += w + space
    
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out

st.divider()
if st.button("🚀 生成 PPT", type="primary", use_container_width=True):
    if st.session_state.images:
        with st.spinner("正在排版生成中..."):
            st.session_state.ppt_bytes = generate_ppt(st.session_state.images)
        st.success("PPT 生成成功！")

if st.session_state.ppt_bytes:
    st.download_button(
        "📂 下载 PPT 文件", 
        data=st.session_state.ppt_bytes, 
        file_name=f"ppt_export_{int(time.time())}.pptx", 
        use_container_width=True
    )
