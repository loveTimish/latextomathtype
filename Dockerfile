FROM eclipse-temurin:21-jre-jammy

RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
        dvisvgm \
        texlive-fonts-recommended \
        texlive-latex-base \
        texlive-latex-extra \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY target/paper-to-word-1.0.0.jar /app/paper-to-word.jar

RUN mkdir -p /var/cache/latextomathtype/formula-render

EXPOSE 8081

ENV JAVA_OPTS="-Dmathtype.windows.enabled=false -Dpaperword.render.cache.enabled=true -Dpaperword.render.cache.dir=/var/cache/latextomathtype/formula-render"

ENTRYPOINT ["sh", "-c", "java $JAVA_OPTS -jar /app/paper-to-word.jar"]
