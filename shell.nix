{ nixpkgs ? import ./pinned.nix {} }:
let
  inherit (nixpkgs) pkgs;

  haskellPackages = pkgs.haskellPackages.override {
    overrides = self: super: {
      binary-search = super.callCabal2nix "binary-search" (builtins.fetchGit {
        url = "git@github.com:piq9117/binary-search.git";
        rev = "acfb1e43fcfaaa364426aa74f40d725ec6f57ffe";
      }){};
    };
  };

  project = haskellPackages.callPackage ./default.nix {};
in
pkgs.stdenv.mkDerivation {
  name = "xlsx";
  buildInputs = project.env.nativeBuildInputs ++ [
    haskellPackages.cabal-install
    haskellPackages.ghc
    haskellPackages.ghcid
  ];
}
