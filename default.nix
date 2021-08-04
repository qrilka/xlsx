let
  rev = "07ca3a021f05d6ff46bbd03c418b418abb781279"; # first 21.05 release
  url = "https://github.com/NixOS/nixpkgs/archive/${rev}.tar.gz";
  compiler = "ghc884";
  pkgs = import (builtins.fetchTarball url) {
    config = {
      packageOverrides = pkgs_super: {
        haskell = pkgs_super.haskell // {
          packages = pkgs_super.haskell.packages // {
            "${compiler}" = pkgs_super.haskell.packages."${compiler}".override {
              overrides = self: super: {
                mkDerivation = args: super.mkDerivation (args // {
                  enableLibraryProfiling = true;
                });
                # libxml-sax = pkgs_super.haskell.lib.dontHaddock (pkgs_super.haskell.lib.dontCheck (
                #   self.callCabal2nix "libxml-sax" ../libxml-sax {}));
              };
            };
          };
        };
      };
    };
  };

  hpkgs = pkgs.haskell.packages."${compiler}";
in pkgs.haskell.lib.dontCheck (hpkgs.callCabal2nix "xlsx" ./. {})
