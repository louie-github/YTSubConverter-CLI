#!/bin/sh

cd YTSubConverter.CLI
mv Program.cs Program.cs.old
sed '/#define WINDOWS/d' Program.cs.old > Program.cs
rm Program.cs.old
