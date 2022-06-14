FROM crystallang/crystal:1.4

RUN apt-get update && apt-get install -y gettext autoconf-archive wget

WORKDIR /workdir

RUN wget https://github.com/libxls/libxls/releases/download/v1.6.2/libxls-1.6.2.tar.gz && \
  tar -xvf libxls-1.6.2.tar.gz && \
  cd libxls-1.6.2 && \
  ./configure && \
  make && \
  make install && \
  cd .. && \
  rm -rf libxls-*

ENV LD_LIBRARY_PATH=/usr/local/lib
