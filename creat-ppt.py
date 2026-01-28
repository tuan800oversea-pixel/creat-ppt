from io import BytesIO
import streamlit as st
from PIL import Image
import uuid
import hashlib
import tempfile
import os
import time

# ================= 页面配置 =================
st.set_page_config(layout="wide")
st.title("多图片自动生成 PPT")

# ================= 初始化 =================
if "images" not in st.session_state:
    st.session_state.images = []

if "processed_ids" not in st.session_state:
    st.session_state.processed_ids = set()

if "page" not in st.session_state:
    st.session_state.page = 1

if "ppt_bytes" not in st.session_state:
    st.session_state.ppt_bytes = None

TMP_DIR = tempfile.gettempdir()

# ================= 上传图片 =================
uploaded_files = st.file_uploader(
    "上传图片（支持批量，建议 ≤500 张，上传图片时点击一张，然后在键盘上同时按住crtl+a，即可全选图片）",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

# 检查文件大小是否超出限制
max_size_mb = 1024  # 1GB = 1024MB
if uploaded_files:
    for file in uploaded_files:
        if len(file.getvalue()) > max_size_mb * 1024 * 1024:
            st.error(f"文件 {file.name} 超过最大上传大小 1GB，请重新选择小于 1GB 的文件")
        else:
            st.success(f"文件 {file.name} 上传成功")

    for idx, file in enumerate(uploaded_files):
        file_bytes = file.read()
        file_hash = hashlib.md5(file_bytes).hexdigest()
        file_id = f"{file.name}_{file_hash}"

        if file_id in st.session_state.processed_ids:
            continue

        try:
            img = Image.open(BytesIO(file_bytes)).convert("RGB")
            w, h = img.size
            ratio = w / h

            thumb = img.copy()
            thumb.thumbnail((260, 260))  # 创建缩略图
            uid = str(uuid.uuid4())
            thumb_path = os.path.join(TMP_DIR, f"{uid}.png")
            thumb.save(thumb_path, "PNG")

            st.session_state.images.append({
                "uid": uid,
                "name": file.name,
                "bytes": file_bytes,
                "thumb_path": thumb_path,
                "ratio": ratio
            })

            st.session_state.processed_ids.add(file_id)

        except Exception as e:
            st.error(f"{file.name} 读取失败：{e}")

        progress.progress((idx + 1) / total)

    progress.empty()
    info.empty()

# ================= 分页参数 =================
IMAGES_PER_PAGE = 40
IMAGES_PER_ROW = 10
THUMB_HEIGHT_MM = 40  # 固定图片高度为40mm
MM_TO_PIXELS = 3.77953  # 1 mm = 3.77953 pixels, this conversion is used for streamlit image sizing

total_images = len(st.session_state.images)
total_pages = max(1, (total_images + IMAGES_PER_PAGE - 1) // IMAGES_PER_PAGE)
st.session_state.page = min(st.session_state.page, total_pages)

start = (st.session_state.page - 1) * IMAGES_PER_PAGE
end = start + IMAGES_PER_PAGE
page_images = st.session_state.images[start:end]

# ================= 控制函数 =================
def prev_page():
    if st.session_state.page > 1:
        st.session_state.page -= 1
        st.rerun()  # 使用 st.rerun() 来重新加载页面

def next_page():
    if st.session_state.page < total_pages:
        st.session_state.page += 1
        st.rerun()  # 使用 st.rerun() 来重新加载页面

# ================= 控制区 =================
st.subheader("图片预览")

# 显示总图片数
st.write(f"共 {total_images} 张图片")

# ================= 图片展示 =================
for i in range(0, len(page_images), IMAGES_PER_ROW):
    cols = st.columns(IMAGES_PER_ROW)
    for col, img in zip(cols, page_images[i:i + IMAGES_PER_ROW]):
        with col:
            # 使用PIL调整图片的高度，宽度按比例缩放
            pil_img = Image.open(img["thumb_path"])

            # 固定高度为40mm，宽度按比例计算
            # 40mm = 40 * 3.77953 pixels
            fixed_height_pixels = int(THUMB_HEIGHT_MM * MM_TO_PIXELS)

            # 计算宽度
            width = int(fixed_height_pixels * img["ratio"])

            # 如果图片的高度小于40mm，则放大至40mm
            pil_img = pil_img.resize((width, fixed_height_pixels))  # 调整图片大小

            # 将调整后的图片传递给 st.image
            img_bytes = BytesIO()
            pil_img.save(img_bytes, format="PNG")
            img_bytes.seek(0)

            # 使用 st.image 展示调整后的图片
            st.image(img_bytes, width=width)  # 使用宽度控制图片的大小

# ================= 分页 =================
cp, cn, ct = st.columns([1, 1, 6])

with cp:
    if st.session_state.page > 1:
        st.button("上一页", on_click=prev_page)

with cn:
    if st.session_state.page < total_pages:
        st.button("下一页", on_click=next_page)

# ================= PPT 生成 =================
def generate_ppt(images, callback=None):
    from pptx import Presentation
    from pptx.util import Inches, Mm
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    left_margin = Mm(0)       # ⚡ 左对齐
    top_margin = Mm(10)    # 每页ppt顶部间隔10mm
    spacing = Mm(2.5)           # 图片之间的水平间距（以及换行后的垂直间距）
    fixed_height = Mm(40)     # 图片固定高度
    max_y = prs.slide_height - top_margin

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    x = left_margin
    y = top_margin

    total = len(images)
    start_time = time.time()

    for idx, img in enumerate(images):
        width = fixed_height * img["ratio"]

        # 换行
        if x + width > prs.slide_width:
            x = left_margin
            y += fixed_height + spacing

        # 超出页面高度 → 新建一页
        if y + fixed_height > max_y:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            x = left_margin
            y = top_margin

        slide.shapes.add_picture(
            BytesIO(img["bytes"]),
            x, y,
            height=fixed_height
        )

        x += width + spacing

        if callback and idx % 5 == 0:
            elapsed = time.time() - start_time
            avg = elapsed / (idx + 1)
            callback(idx / total * 0.95, avg * (total - idx - 1))

    output = BytesIO()
    prs.save(output)
    output.seek(0)

    if callback:
        callback(1.0, 0)

    return output

# ================= 生成 & 下载 =================
st.divider()

if st.button("生成 PPT"):
    # 直接选择所有图片
    selected = st.session_state.images

    if not selected:
        st.warning("请至少选择一张图片")
    else:
        bar = st.progress(0.0)
        text = st.empty()

        def cb(p, t):
            bar.progress(p)
            text.text(f"预计剩余时间：{t:.1f}s")

        st.session_state.ppt_bytes = generate_ppt(selected, cb)

        bar.empty()
        text.empty()

        st.download_button(
            "下载 PPT",
            data=st.session_state.ppt_bytes,
            file_name="images.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )


