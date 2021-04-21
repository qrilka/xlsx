{
  description = "xlsx";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs";
    flake-utils.url = "github:numtide/flake-utils/5466c5bbece17adaab2d82fae80b46e807611bf3";
    gitignore = {
      url = "github:hercules-ci/gitignore.nix/211907489e9f198594c0eb0ca9256a1949c9d412";
      flake = false;
    };
  };

  outputs = { self, nixpkgs, flake-utils, gitignore }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        packageName = "xlsx";
        compiler = "ghc884";

        pkgs = nixpkgs.legacyPackages.${system};
        jailbreakUnbreak = pkg:
          pkgs.haskell.lib.doJailbreak (pkg.overrideAttrs (_: { meta = { }; }));

        gitignore = pkgs.nix-gitignore.gitignoreSourcePure [ ./.gitignore ];

      in {
        packages.${packageName} =
          pkgs.haskell.packages.${compiler}.callCabal2nix packageName (gitignore ./.) rec { };

        defaultPackage = self.packages.${system}.${packageName};

        devShell = pkgs.mkShell {
          buildInputs = with pkgs; [
            haskell-language-server
            stylish-haskell
            ghcid
            zlib 
            cabal-install
          ];
          inputsFrom = builtins.attrValues self.packages.${system};
          # shellHook = "cabal v2-repl";
        };
      });
}