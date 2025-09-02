{-# LANGUAGE CPP #-}
{-# LANGUAGE RankNTypes #-}

module Codec.Xlsx.LensCompat 
  ( module L
  , Prism
  , Prism'
  , prism
  , (<>=)
  ) where

#ifdef USE_MICROLENS

import Lens.Micro as L
import Lens.Micro.GHC ()
import Lens.Micro.Mtl as L
import Lens.Micro.Platform as L
import Data.Profunctor.Choice
import Data.Profunctor (dimap)
import Data.Traversable.WithIndex as L
import Lens.Micro.Internal as L
import Control.Monad.State.Class (MonadState, modify)

-- Since micro-lens denies the existence of prisms,
-- I copied over the definitions from Control.Lens for the prism
-- function.
type Prism s t a b = forall p f. (Choice p, Applicative f) => p a (f b) -> p s (f t)
type Prism' s a = Prism s s a a

prism :: (b -> t) -> (s -> Either t a) -> Prism s t a b
prism bt seta = dimap seta (either pure (fmap bt)) . right'

(<>=) :: (MonadState s m, Monoid a) => ASetter' s a -> a -> m ()
l <>= a = modify (l <>~ a)

#else

import Control.Lens as L

#endif
