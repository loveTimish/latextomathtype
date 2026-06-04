#!/usr/bin/env sh
set -eu

PORT="${PORT:-18082}"
CACHE_DIR="${CACHE_DIR:-target/linux-smoke-cache}"
JAR="${JAR:-target/paper-to-word-1.0.0.jar}"

if [ ! -f "$JAR" ]; then
  echo "missing jar: $JAR" >&2
  echo "build it first: ./.mvn/apache-maven-3.9.12/bin/mvn -DskipTests package" >&2
  exit 1
fi

mkdir -p "$CACHE_DIR"
test -w "$CACHE_DIR"

java \
  -Dserver.port="$PORT" \
  -Dmathtype.windows.enabled=false \
  -Dpaperword.render.cache.enabled=true \
  -Dpaperword.render.cache.dir="$CACHE_DIR" \
  -jar "$JAR" >/tmp/latextomathtype-linux-smoke.log 2>&1 &

PID="$!"
cleanup() {
  kill "$PID" >/dev/null 2>&1 || true
}
trap cleanup EXIT INT TERM

i=0
while [ "$i" -lt 30 ]; do
  if command -v curl >/dev/null 2>&1; then
    if curl -fsS "http://127.0.0.1:$PORT/api/export/health" | grep -q "paper-to-word service is running"; then
      echo "linux smoke ok: health endpoint reachable, cache=$CACHE_DIR"
      exit 0
    fi
  elif command -v wget >/dev/null 2>&1; then
    if wget -qO- "http://127.0.0.1:$PORT/api/export/health" | grep -q "paper-to-word service is running"; then
      echo "linux smoke ok: health endpoint reachable, cache=$CACHE_DIR"
      exit 0
    fi
  else
    echo "curl or wget is required for health check" >&2
    exit 1
  fi
  i=$((i + 1))
  sleep 1
done

echo "linux smoke failed; service log:" >&2
cat /tmp/latextomathtype-linux-smoke.log >&2
exit 1
