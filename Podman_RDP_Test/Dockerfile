# Dockerfile para imagem RDP/VNC com LXDE
FROM ubuntu:22.04

# Instalação otimizada de pacotes
RUN apt-get update &&
  DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
    xrdp \
    xorgxrdp \
    x11vnc \
    novnc \
    websockify \
    lxde-core \
    firefox \
    lxsession \
    tigervnc-standalone-server \
    sudo \
    passwd &&
  apt-get clean &&
  rm -rf /var/lib/apt/lists/*

# Configurações essenciais
RUN echo "startlxde" >/etc/xrdp/startwm.sh &&
  adduser xrdp ssl-cert &&
  echo "lxsession" >/etc/skel/.xsession

# Cria um usuário padrão (user/password)
RUN useradd -m -s /bin/bash user &&
  echo "user:password" | chpasswd &&
  usermod -aG sudo user

# Configura o VNC para usar senha vazia e permitir conexões locais
RUN mkdir -p /home/user/.vnc &&
  echo "" | vncpasswd -f >/home/user/.vnc/passwd &&
  chmod 600 /home/user/.vnc/passwd &&
  chown -R user:user /home/user/.vnc

# Expõe as portas padrão
EXPOSE 3389 6080

# Comando de inicialização otimizado (corrige problema do VNC)
CMD ["sh", "-c", "service xrdp restart && x11vnc -display :0 -nopw -listen 0.0.0.0 -xkb -forever & websockify --web=/usr/share/novnc 6080 localhost:5900"]
