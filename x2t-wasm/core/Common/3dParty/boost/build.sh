#!/bin/bash
SCRIPTPATH="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"

os=$(uname -s)
platform=""

case "$os" in
  Linux*)   platform="linux" ;;
  Darwin*)  platform="mac" ;; 
  *)        exit ;;
esac

architecture="x86"
arch=""

case "$architecture" in
  x86_64*)  arch="_64" ;;
  *)        arch="_32" ;;
esac

outputdir="$SCRIPTPATH/boost_1_58_0/build/$platform$arch"
echo $outputdir

cd "$SCRIPTPATH/boost_1_58_0"

folder="build/$platform$arch"
if [ ! -d $folder ]; then
  ./bootstrap.sh --with-toolset=gcc --with-libraries=filesystem,system,date_time,regex

  ## Custom gcc.jam to force correct ar and ranlib which fail under emscripten
  cp ../gcc.jam ./tools/build/src/tools/gcc.jam
  mv /usr/bin/rc /usr/bin/linux-rc
  ln -s /usr/bin/linux-rc /usr/bin/rc
  update-alternatives --install /usr/bin/gcc gcc /emsdk/upstream/emscripten/emcc 20
  update-alternatives --install /usr/bin/g++ g++ /emsdk/upstream/emscripten/em++ 20
  update-alternatives --install /usr/bin/clang clang /emsdk/upstream/bin/clang 20
  update-alternatives --install /usr/bin/clang++ clang++ /emsdk/upstream/bin/clang++ 20
  update-alternatives --install /usr/bin/ar ar /emsdk/upstream/bin/llvm-ar 20
  update-alternatives --install /usr/bin/rc rc /emsdk/upstream/bin/llvm-rc 20

   stage="stage"
   if [ -d $stage ]; then
     rm -R $stage
   fi

  mkdir -p "$folder"
  mkdir -p "$folder/static"
  mkdir -p "$folder/static_fpic"
  mkdir -p "$folder/shared"

  ./b2 --clean
  ./bjam link=static --debug-configuration
  cp stage/lib/* "$folder/static/"
 
#  ./b2 --clean
#  ./bjam link=static cxxflags=-fPIC
#  cp stage/lib/* "$folder/static_fpic/"

#  ./b2 --clean
#  ./bjam link=shared
#  cp stage/lib/* "$folder/shared/"
fi
