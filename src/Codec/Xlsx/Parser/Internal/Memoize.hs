{-# LANGUAGE GeneralizedNewtypeDeriving #-}
{-# LANGUAGE FlexibleContexts #-}
{-# LANGUAGE TypeApplications #-}

-- | I rewrote: https://hackage.haskell.org/package/unliftio-0.2.20/docs/src/UnliftIO.Memoize.html#Memoized
-- for monad trans basecontrol
-- we don't need a generic `m` anyway. it's good enough in base IO.
module Codec.Xlsx.Parser.Internal.Memoize
  ( Memoized
  , runMemoized
  , memoizeRef
  ) where

import Control.Applicative as A
import Control.Monad (join)
import Control.Monad.IO.Class
import Data.IORef
import Control.Exception

-- | A \"run once\" value, with results saved. Extract the value with
-- 'runMemoized'. For single-threaded usage, you can use 'memoizeRef' to
-- create a value. If you need guarantees that only one thread will run the
-- action at a time, use 'memoizeMVar'.
--
-- Note that this type provides a 'Show' instance for convenience, but not
-- useful information can be provided.
--
-- @since 0.2.8.0
newtype Memoized a = Memoized (IO a)
  deriving (Functor, A.Applicative, Monad)
instance Show (Memoized a) where
  show _ = "<<Memoized>>"

-- | Extract a value from a 'Memoized', running an action if no cached value is
-- available.
--
-- @since 0.2.8.0
runMemoized :: MonadIO m => Memoized a -> m a
runMemoized (Memoized m) = liftIO m
{-# INLINE runMemoized #-}

-- | Create a new 'Memoized' value using an 'IORef' under the surface. Note that
-- the action may be run in multiple threads simultaneously, so this may not be
-- thread safe (depending on the underlying action). Consider using
-- 'memoizeMVar'.
--
-- @since 0.2.8.0
memoizeRef :: IO a -> IO (Memoized a)
memoizeRef action = do
  ref <- newIORef Nothing
  pure $ Memoized $ do
    mres <- readIORef ref
    res <-
      case mres of
        Just res -> pure res
        Nothing -> do
          res <- try @SomeException action
          writeIORef ref $ Just res
          pure res
    either throwIO pure res
