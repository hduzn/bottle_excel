# Use an official Python runtime as an image
FROM python:3.9

# Sets the working directory for following COPY and CMD instructions
# Notice we haven’t created a directory by this name - this instruction
# creates a directory with this name if it doesn’t exist
WORKDIR /app

# copy code
COPY . /app
# COPY requirements.txt /app

# RUN python -m pip install --upgrade pip
# RUN pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host=files.pythonhosted.org --no-cache-dir -r requirements.txt
RUN pip3 install -i https://pypi.tuna.tsinghua.edu.cn/simple --upgrade pip 
RUN pip3 install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt 

# The EXPOSE instruction indicates the ports on which a container
# 我这里app.py中run的port=9881，所以开的是这个端口，你的看你自己的设置
EXPOSE 9881

# Run app.py when the container launches
# COPY app.py /app
CMD ["python3", "app.py"]