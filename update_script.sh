#/bin/bash

pm2 stop main
git reset --hard
git pull
pm2 start main