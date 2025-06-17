FROM ubuntu:22.04

# Comandos separados (mais compatível)
RUN apt-get update
RUN DEBIAN_FRONTEND=noninteractive apt-get install -y xrdp xorgxrdp
RUN DEBIAN_FRONTEND=noninteractive apt-get install -y x11vnc novnc websockify
RUN DEBIAN_FRONTEND=noninteractive apt-get install -y lxde-core firefox
RUN apt-get clean

RUN echo "startlxde" >/etc/xrdp/startwm.sh
RUN adduser xrdp ssl-cert
RUN echo "xfce4-session" >/etc/skel/.xsession

EXPOSE 3389 6080
CMD ["sh", "-c", "service xrdp start && websockify --web=/usr/share/novnc 6080 localhost:5900"]
