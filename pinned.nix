let
  pkgs = import <nixpkgs> {};
in
import (
  pkgs.fetchFromGitHub {
    owner = "nixos";
    repo = "nixpkgs";
    rev = "215337d484db4d981bd88ecf241fae2a263eacf2";
    sha256 = "1qswdihjhsvwp6yzzhi7vkk1nijq2hvghj2axd9vjzs0cv18y3gh";
  }
)
# https://github.com/NixOS/nixpkgs/archive/215337d484db4d981bd88ecf241fae2a263eacf2.tar.gz
