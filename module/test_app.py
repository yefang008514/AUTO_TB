###############################按钮形式#########################################

# import streamlit as st


# # 页面标题
# st.set_page_config(page_title="按钮样式调整", page_icon="🎨", layout="wide")
# st.title("自定义按钮样式示例")

# # 自定义按钮的 CSS 样式
# button_style = """
#     <style>
#     div.stButton > button {
#         width: 200px; /* 统一按钮宽度 */
#         height: 50px; /* 可选：统一高度 */
#         margin: 5px auto; /* 设置按钮之间的间距 */
#         font-size: 16px; /* 字体大小 */
#     }
#     </style>
# """
# st.markdown(button_style, unsafe_allow_html=True)

# # # 页面标题和图片
# # st.set_page_config(page_title="多页面示例", page_icon="📄", layout="wide")
# # st.image("https://via.placeholder.com/1500x300", caption="File Q&A with Anthropic")  # 替换为你的图片路径
# # st.title("文件问答与其他功能")

# # 使用 Streamlit 的 Session State 管理按钮状态
# if "current_page" not in st.session_state:
#     st.session_state.current_page = "首页"

# # 定义页面切换函数
# def change_page(page_name):
#     st.session_state.current_page = page_name

# # 侧边栏按钮布局
# st.sidebar.title("导航")
# with st.sidebar:
#     if st.button("首页"):
#         change_page("首页")
#     if st.button("文件问答"):
#         change_page("文件问答")
#     if st.button("搜索聊天"):
#         change_page("搜索聊天")
#     if st.button("Langchain 快速开始"):
#         change_page("Langchain 快速开始")
#     if st.button("Langchain PromptTemplate"):
#         change_page("Langchain PromptTemplate")
#     if st.button("用户反馈聊天"):
#         change_page("用户反馈聊天")

# # 页面内容
# if st.session_state.current_page == "首页":
#     st.header("🏠 首页")
#     st.write("欢迎访问首页！您可以在此添加一些项目概览或介绍内容。")

# elif st.session_state.current_page == "文件问答":
#     st.header("📄 文件问答")
#     uploaded_file = st.file_uploader("上传文件 (支持 TXT 或 MD 格式)", type=["txt", "md"])
#     if uploaded_file:
#         content = uploaded_file.read().decode("utf-8")
#         st.text_area("文件内容", content, height=300)
#         st.text_input("问文件什么问题？", placeholder="例如：可以给我一个简短总结吗？")

# elif st.session_state.current_page == "搜索聊天":
#     st.header("🔍 搜索聊天")
#     st.text_input("输入您的搜索问题", placeholder="搜索您感兴趣的内容")
#     st.button("开始搜索")

# elif st.session_state.current_page == "Langchain 快速开始":
#     st.header("🚀 Langchain 快速开始")
#     st.write("这里可以展示 Langchain 的快速入门教程或示例代码。")

# elif st.session_state.current_page == "Langchain PromptTemplate":
#     st.header("🧩 Langchain PromptTemplate")
#     st.write("展示 PromptTemplate 的详细信息或使用示例。")

# elif st.session_state.current_page == "用户反馈聊天":
#     st.header("💬 用户反馈聊天")
#     st.text_input("输入您的反馈内容", placeholder="在此填写您的意见或建议")
#     st.button("提交反馈")

# # 添加版权信息
# st.sidebar.write("---")
# st.sidebar.write("© 2025 您的项目名称")



################################双侧边栏#################################


import streamlit as st

# 页面配置
st.set_page_config(page_title="双侧边栏示例", page_icon="📋", layout="wide")

# 主侧边栏导航
st.sidebar.title("主侧边栏")
main_section = st.sidebar.radio("选择主页面", ["首页", "文件问答", "设置"])

# 页面逻辑
if main_section == "首页":
    st.title("🏠 首页")
    st.write("这是首页内容。")

elif main_section == "文件问答":
    st.title("📄 文件问答")

    # 模拟子侧边栏
    with st.sidebar.expander("子导航"):
        sub_section = st.radio("选择子功能", ["上传文件", "问答历史", "文件设置"])
    
    if sub_section == "上传文件":
        st.header("上传文件")
        uploaded_file = st.file_uploader("上传文件", type=["txt", "md"])
        if uploaded_file:
            content = uploaded_file.read().decode("utf-8")
            st.text_area("文件内容", content, height=300)
    
    elif sub_section == "问答历史":
        st.header("问答历史")
        st.write("这里是问答记录的展示区域。")
    
    elif sub_section == "文件设置":
        st.header("文件设置")
        st.write("可以设置一些与文件相关的选项。")

elif main_section == "设置":
    st.title("⚙ 设置")
    st.write("这是设置页面，您可以在此调整应用参数。")
