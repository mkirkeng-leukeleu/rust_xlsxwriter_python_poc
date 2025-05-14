.PHONY: build
build:
	uv tool run maturin build --release

.PHONY: install
install:
	curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
	. "$HOME/.cargo/env"
	uv tool install maturin
