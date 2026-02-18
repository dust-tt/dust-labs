import sys

from destination_dust.destination import DestinationDust


def main():
    DestinationDust().run(sys.argv[1:])


if __name__ == "__main__":
    main()
