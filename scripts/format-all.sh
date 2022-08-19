#!/bin/bash

# stack install stylish-haskell
# or grab a release at https://github.com/haskell/stylish-haskell/releases

src_dir="$(realpath $(dirname $0))/.."
config_file="${src_dir}/.stylish-haskell.yaml"
dirs="$(echo ${src_dir}/{benchmarks,src,test})"

stylish-haskell \
	--config "$config_file" \
	--inplace \
	--verbose \
	$(find ${dirs} -type f -name '*.hs')
