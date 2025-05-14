.PHONY: build
build:
	uv tool run maturin build --release
	uv pip install /tasks/xlsx-rust/target/wheels/xlsx_rust-0.1.0-cp313-cp313-manylinux_2_34_aarch64.whl --system --force-reinstall
	# supervisorctl restart jupyter

.PHONY: install
install:
	curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
	. "$HOME/.cargo/env"
	uv tool install maturin
