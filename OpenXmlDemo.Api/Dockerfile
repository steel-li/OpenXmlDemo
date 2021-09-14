#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:5.0 AS base
# 图像操作需要安装引用的
FROM steelli98/net.drawing
WORKDIR /app
EXPOSE 80
EXPOSE 443

# 安装fontconfig库，用于word,Pdf导出
RUN apt-get update && apt-get install -y fontconfig
RUN ln -sf /usr/share/zoneinfo/Asia/Shanghai /etc/localtime
RUN echo 'Asia/Shanghai' >/etc/timezone

ENTRYPOINT ["dotnet", "OpenXmlDemo.Api.dll", "-b", "0.0.0.0"]